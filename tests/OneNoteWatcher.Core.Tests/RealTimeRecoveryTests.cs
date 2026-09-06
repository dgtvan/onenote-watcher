using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Health;
using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

/// <summary>
/// Regression tests for the 2026-09-06 report: a real-time upload error on Note / eSim Data was raised
/// correctly, but the content DID reach OneDrive and the error never went away.
/// </summary>
public class EmbeddedErrorCodeTests
{
    // the exact payload shape OneNote emitted: the code lives only inside the free-text Error string
    private const string RealTimeFailure = """
        SendEvent {"EventName":"Office.OneNote.Storage.RealTime.NoteItService",
        "Data.OperationWithError":"Http Patch for Upload Failed",
        "Data.Error":"Win32Error: ErrOutOfSyncWithStore (0xE000002E) tag_4oxx9",
        "Data.SectionId_ResourceId":"36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2",
        "Data.NotebookId_ResourceId":"36B934175DC7E3A4!626"}
        """;

    [Fact]
    public void Code_embedded_in_error_text_is_extracted()
    {
        var ev = SyncEventJson.TryParse(RealTimeFailure.ReplaceLineEndings(""));
        Assert.NotNull(ev);
        Assert.Equal(0xE000002Eu, ev!.ErrorCode);
        Assert.True(ev.IsProblem);
    }

    [Fact]
    public void Extracted_code_gives_the_catalogued_reason_not_a_guess()
    {
        var ev = SyncEventJson.TryParse(RealTimeFailure.ReplaceLineEndings(""))!;
        var ex = ErrorCatalog.Explain(ev.ErrorCode, ev.ErrorDescription, ev.ErrorType);
        // was previously guessed as Network ("could not be reached") purely from the word "Http"
        Assert.Equal(FailureCategory.Conflict, ex.Category);
        Assert.Contains("out of sync", ex.Summary, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("no code here at all")]
    [InlineData("value 0x00000000 means nothing")]
    public void Text_without_a_usable_code_leaves_it_zero(string text)
    {
        var json = $$"""SendEvent {"EventName":"Office.OneNote.Storage.RealTime.NoteItService","Data.Error":"{{text}}"}""";
        var ev = SyncEventJson.TryParse(json);
        Assert.NotNull(ev);
        Assert.Equal(0u, ev!.ErrorCode);
    }
}

public class CloudConfirmationTests
{
    private static GraphSnapshot Server(DateTimeOffset lastModified) => new(
        DateTimeOffset.UtcNow,
        [new GraphNotebook("nb", "Note", lastModified)],
        [new GraphSection("sec", "eSim Data", lastModified, "nb", "Note", null, null)]);

    private static IReadOnlyList<LocalSection> Local(DateTimeOffset newest) =>
        [new LocalSection("Note", "eSim Data", newest)];

    private static OutcomeDetector Baselined(DateTimeOffset t)
    {
        var d = new OutcomeDetector(TimeSpan.FromMinutes(10));
        d.Evaluate(Local(t), Server(t), true, true, t);   // first poll only learns
        return d;
    }

    [Fact]
    public void Server_ahead_of_the_failure_is_proof_the_change_landed()
    {
        var t0 = DateTimeOffset.UtcNow.AddHours(-1);
        var d = Baselined(t0);
        var failedAt = t0.AddMinutes(5);
        var serverAt = failedAt.AddMinutes(2);

        d.Evaluate(Local(serverAt), Server(serverAt), true, true, serverAt.AddSeconds(30));

        Assert.True(d.CloudConfirmedAfter("Note", "eSim Data", failedAt));
    }

    [Fact]
    public void Server_older_than_the_failure_proves_nothing_so_the_error_stands()
    {
        var t0 = DateTimeOffset.UtcNow.AddHours(-1);
        var d = Baselined(t0);
        d.Evaluate(Local(t0), Server(t0), true, true, t0.AddMinutes(1));

        Assert.False(d.CloudConfirmedAfter("Note", "eSim Data", t0.AddMinutes(30)));
    }

    [Fact]
    public void A_section_never_compared_is_never_confirmed()
    {
        var d = Baselined(DateTimeOffset.UtcNow.AddHours(-1));
        Assert.False(d.CloudConfirmedAfter("Note", "Some Other Section", DateTimeOffset.UtcNow.AddDays(-1)));
        Assert.False(d.CloudConfirmedAfter("Note", null, DateTimeOffset.UtcNow.AddDays(-1)));
    }

    [Fact]
    public void Server_behind_the_local_copy_is_not_a_confirmation()
    {
        var t0 = DateTimeOffset.UtcNow.AddHours(-2);
        var d = Baselined(t0);
        // local raced ahead by an hour; the server has NOT caught up
        var localNow = t0.AddHours(1);
        d.Evaluate(Local(localNow), Server(t0.AddMinutes(1)), true, true, localNow);

        Assert.False(d.CloudConfirmedAfter("Note", "eSim Data", t0));
    }
}

/// <summary>
/// Reported 2026-09-06: after a 3 h sleep the watcher said "The Microsoft Graph check is not running"
/// and told the user to sign in — while they WERE signed in and no sign-in button was even shown.
/// </summary>
public class GraphStaleDiagnosisTests
{
    [Fact]
    public void A_stalled_poll_is_not_diagnosed_as_a_sign_in_problem()
    {
        var now = DateTimeOffset.UtcNow;
        var issue = HealthIssues.GraphStale(TimeSpan.FromHours(3), now.AddHours(-3), now);

        Assert.NotEqual(IssueKind.AuthRequired, issue.Kind);
        Assert.NotEqual(FailureCategory.Permission, issue.Category);
        Assert.DoesNotContain("Sign in", issue.Recommendation, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Check now", issue.Recommendation);
    }

    [Fact]
    public void An_expired_sign_in_still_is_an_auth_problem_with_the_sign_in_action()
    {
        var issue = HealthIssues.GraphTokenExpired(DateTimeOffset.UtcNow);

        Assert.Equal(IssueKind.AuthRequired, issue.Kind);
        Assert.Contains("Sign in", issue.Recommendation);
    }
}

public class AwakeClockTests
{
    [Fact]
    public void The_os_counter_is_available_on_this_platform()
    {
        Assert.True(AwakeClock.Available);
    }

    [Fact]
    public void Elapsed_never_exceeds_the_time_we_were_awake()
    {
        var stamp = AwakeClock.Stamp();
        var wallSince = DateTimeOffset.UtcNow.AddHours(-3);   // as if the machine slept for three hours

        var elapsed = AwakeClock.Elapsed(wallSince, DateTimeOffset.UtcNow, stamp);

        Assert.True(elapsed < TimeSpan.FromMinutes(1), $"reported {elapsed} of running time after no running time");
    }

    [Fact]
    public void Elapsed_never_exceeds_the_wall_clock_either()
    {
        var now = DateTimeOffset.UtcNow;
        Assert.True(AwakeClock.Elapsed(now, now, AwakeClock.Stamp() - TimeSpan.FromHours(5)) < TimeSpan.FromSeconds(1));
    }
}
