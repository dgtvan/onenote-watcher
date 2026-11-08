using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher.Core.Tests;

/// <summary>
/// Regressions from the 2026-11-08/09 incident, where OneNote's sign-in was rejected for two days.
/// OneNote reported it correctly; the watcher showed 17 issues that named neither the cause nor the
/// fix, so the signal was there and unreadable. Payloads below are real, from
/// <c>C:\ProgramData\OneNoteWatcher\logs\collector-2026-11-0[89].log</c>.
/// </summary>
public class AuthIncidentRegressionTests
{
    private static SyncEvent Parse(string json) => SyncEventJson.TryParse("SendEvent " + json)!;

    // ---------- the error description that names the fault is nested, and was being dropped ----------

    private const string PageSyncSessionAccessDenied = """
        {"EventName":"Office.OneNote.Storage.PageSyncSession","Time":"2026-11-07T08:37:47.497Z",
         "Data.ActivePageGOID":"{852721A5-6CDB-41ED-AD20-FA5608E6BCCA}{14}",
         "Data.ErrorState_TimeInHierarchySyncErrorState":0,
         "Data.ErrorState_TimeInActiveSyncErrorState":1534,
         "Data.ErrorState_Errors_SectionObjectSpace_InboundError_Error_Tag":"4pfvq",
         "Data.ErrorState_Errors_SectionObjectSpace_InboundError_Error_Type":"Win32Error",
         "Data.ErrorState_Errors_SectionObjectSpace_InboundError_Error_Code":3758096494,
         "Data.ErrorState_Errors_SectionObjectSpace_InboundError_Error_Descript":"ErrEndpointAccessDenied",
         "Data.ErrorState_Errors_Page_InboundError_Error_Type":"Win32Error",
         "Data.ErrorState_Errors_Page_InboundError_Error_Code":3758096494,
         "Data.ErrorState_Errors_Page_InboundError_Error_Description":"ErrEndpointAccessDenied"}
        """;

    [Fact]
    public void Nested_error_description_is_read_so_the_issue_names_the_fault()
    {
        var e = Parse(PageSyncSessionAccessDenied);

        Assert.Equal(0xE000006Eu, e.ErrorCode);
        Assert.Equal("ErrEndpointAccessDenied", e.ErrorDescription);
        Assert.Equal("Win32Error", e.ErrorType);
        // time in an error state is still hard evidence, so the outcome is unchanged
        Assert.Equal(SyncOutcome.SuspectedFailure, e.Outcome);

        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Equal(FailureCategory.Permission, issue.Category);
        Assert.Equal("ErrEndpointAccessDenied", issue.Message);
        Assert.Contains("File", issue.Recommendation);        // "File -> Account", i.e. sign in again
    }

    [Fact]
    public void Truncated_Error_Descript_spelling_is_read_too()
    {
        // Office truncates long field names; the same payload carried both spellings.
        var e = Parse("""
            {"EventName":"Office.OneNote.Storage.RealTime.Whatever","Time":"2026-11-08T17:10:52Z",
             "Data.Some_Nested_Error_Descript":"ErrCobaltAuthenticationFailure"}
            """);
        Assert.Equal("ErrCobaltAuthenticationFailure", e.ErrorDescription);
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
    }

    // ---------- a field named "...ErrorCode" is not automatically an error code ----------

    private const string GetSharePointIds = """
        {"EventName":"Office.OneNote.GetSharePointIdsForDocument","Time":"2026-11-08T16:11:33Z",
         "Data.ErrorCode":2,"Data.HttpStatusCode":401,
         "Data.ErrorMessage":"ohgI+BySYz4aMY1ToXtlRHLIcGWJN4MTLJWqBtcs9XA=","Data.InnerErrorCode":14}
        """;

    [Fact]
    public void Internal_enum_is_not_promoted_to_a_sync_error_code()
    {
        var e = Parse(GetSharePointIds);

        Assert.Equal(0u, e.ErrorCode);           // NOT 0x00000002
        Assert.Null(e.ErrorCodeHex);
        // the evidence is kept, it is just not the headline
        Assert.Equal("2", e.ErrorFields["Data.ErrorCode"]);
        Assert.Equal("401", e.ErrorFields["Data.HttpStatusCode"]);
    }

    [Fact]
    public void Hashed_ErrorMessage_never_becomes_the_message_shown_to_the_user()
    {
        var e = Parse(GetSharePointIds);

        Assert.DoesNotContain("ohgI", e.ErrorDescription ?? "");
        Assert.Equal("HTTP 401 Unauthorized", e.ErrorDescription);
        // still a failure, still recorded in full
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        Assert.Contains("ohgI", e.ErrorFields["Data.ErrorMessage"]);

        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Equal(FailureCategory.Permission, issue.Category);   // not Network: 401 means sign in
        Assert.DoesNotContain("ohgI", issue.Summary);
        Assert.Contains("ohgI", issue.TechnicalDetail);
    }

    [Fact]
    public void A_genuine_hresult_is_still_promoted()
    {
        var e = Parse("""
            {"EventName":"Office.OneNote.Storage.NotebookSyncResult","Time":"2026-11-08T16:11:30Z",
             "Data.CachedError_Code":3758096420,"Data.NotebookErrorCode":3758096420}
            """);
        Assert.Equal(0xE0000024u, e.ErrorCode);
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        Assert.Equal(FailureCategory.Permission, SyncEventMapper.ToIssue(e, "etw")!.Category);
    }

    // ---------- SyncHealth is a counter, not an outcome ----------

    [Fact]
    public void SyncHealth_counter_is_not_an_issue()
    {
        var e = Parse("""
            {"EventName":"Office.OneNote.Storage.SyncHealth","Time":"2026-11-07T08:37:45Z",
             "Data.Tag":"4pfvq","Data.Service":0}
            """);
        Assert.Equal(SyncOutcome.Diagnostic, e.Outcome);
        Assert.False(e.IsProblem);
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
    }

    [Fact]
    public void SyncHealth_carrying_an_actual_error_still_surfaces()
    {
        // listed in BenignUnlessError, not DiagnosticEvents: silence is only for a quiet payload
        var e = Parse("""
            {"EventName":"Office.OneNote.Storage.SyncHealth","Time":"2026-11-07T08:37:45Z",
             "Data.Tag":"4pfvq","Data.Error_Code":3758096494,"Data.Error_Description":"ErrEndpointAccessDenied"}
            """);
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        Assert.True(e.IsProblem);
    }

    // ---------- the codes seen throughout the incident now carry an explanation ----------

    [Theory]
    [InlineData(0xE0000024u)]   // authentication (Microsoft's own article)
    [InlineData(0xE000006Eu)]   // ErrEndpointAccessDenied
    [InlineData(0xE000006Fu)]   // ErrCobaltAuthenticationFailure
    public void Incident_codes_are_catalogued_as_permission_problems(uint code)
    {
        var ex = ErrorCatalog.Explain(code, null);
        Assert.Equal(FailureCategory.Permission, ex.Category);
        Assert.DoesNotContain("unrecognised", ex.Summary);
    }

    [Fact]
    public void Outbound_blocked_code_is_catalogued_as_a_knock_on_conflict()
    {
        var ex = ErrorCatalog.Explain(0xE0000057u, null);
        Assert.Equal(FailureCategory.Conflict, ex.Category);
        Assert.True(ex.LikelyTransient);
    }

    // ---------- the "local is ahead" clock must not outlive the condition ----------

    private static GraphSnapshot Snapshot(DateTimeOffset server) => new(
        DateTimeOffset.UtcNow, [],
        [new GraphSection("sec-1", "Videos", server, "nb-1", "Note", null, null)]);

    [Fact]
    public void Ahead_clock_is_cleared_when_the_local_timestamp_falls_back()
    {
        var t0 = new DateTimeOffset(2026, 11, 8, 16, 0, 0, TimeSpan.Zero);
        var server = new DateTimeOffset(2026, 11, 3, 2, 9, 51, TimeSpan.Zero);
        var d = new OutcomeDetector(TimeSpan.FromMinutes(10));

        LocalSection At(DateTimeOffset t) => new("Note", "Videos", t);

        // baseline, then the local stamp jumps ahead and stays there past the grace period
        Assert.Empty(d.Evaluate([At(server)], Snapshot(server), true, true, t0));
        Assert.Empty(d.Evaluate([At(t0)], Snapshot(server), true, true, t0));
        var stuck = d.Evaluate([At(t0)], Snapshot(server), true, true, t0.AddMinutes(45));
        Assert.Equal(IssueKind.UploadStuck, Assert.Single(stuck).Kind);

        // the index re-stamps the section back to the baseline: nothing is pending any more
        Assert.Empty(d.Evaluate([At(server)], Snapshot(server), true, true, t0.AddMinutes(50)));
        Assert.Contains("Videos", Assert.Single(d.LastResolved));

        // ...and the clock really is gone: a fresh local change gets the full grace period again,
        // instead of alerting instantly on the hours already banked.
        var t1 = t0.AddMinutes(60);
        Assert.Empty(d.Evaluate([At(t1)], Snapshot(server), true, true, t1));
        Assert.Empty(d.Evaluate([At(t1)], Snapshot(server), true, true, t1.AddMinutes(5)));
        Assert.Single(d.Evaluate([At(t1)], Snapshot(server), true, true, t1.AddMinutes(11)));
    }

    // ---------- a lost race with the collector's file replace is not "the collector is down" ----------

    [Fact]
    public async Task Status_load_retries_past_a_transient_replace_window()
    {
        var dir = Directory.CreateTempSubdirectory().FullName;
        try
        {
            var path = Path.Combine(dir, "status.json");
            new WatcherStatus { CollectorRunning = true }.Save(path);

            // hold the file the way a mid-flight File.Move does: no sharing at all
            var block = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            var unblock = Task.Run(async () => { await Task.Delay(40); block.Dispose(); });

            var loaded = await Task.Run(() => WatcherStatus.Load(path));
            await unblock;

            Assert.NotNull(loaded);
            Assert.True(loaded!.CollectorRunning);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void Status_load_still_returns_null_when_the_file_is_really_gone()
    {
        var dir = Directory.CreateTempSubdirectory().FullName;
        try { Assert.Null(WatcherStatus.Load(Path.Combine(dir, "status.json"))); }
        finally { Directory.Delete(dir, true); }
    }
}
