using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;
using OneNoteWatcher.Core.Rules;

namespace OneNoteWatcher.Core.Tests;

/// <summary>
/// The fail-closed contract (docs/fail-closed.md): only an explicit positive success signal
/// counts as success; anything else must surface. These tests are the guard against silently missing
/// a rare sync failure.
/// </summary>
public class FailClosedTests
{
    private static SyncEvent Parse(string json) => SyncEventJson.TryParse("SendEvent " + json)!;

    // ---------- unknown / future events are never dropped ----------

    [Fact]
    public void Unknown_storage_event_with_no_fields_is_Unknown_not_success()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SomeFutureThing","Time":"2026-09-06T10:00:00Z"}""");
        Assert.Equal(SyncEventKind.UnrecognisedStorage, e.Kind);
        Assert.Equal(SyncOutcome.Unknown, e.Outcome);
        Assert.True(e.IsProblem);
        Assert.False(e.IsSyncActivity);

        var issue = SyncEventMapper.ToIssue(e, "etw")!;
                Assert.Contains("does not recognise", issue.Summary);
        Assert.Contains("event=Office.OneNote.Storage.SomeFutureThing", issue.TechnicalDetail);
        Assert.NotNull(issue.ExpiresUtc);          // self-clears, so it cannot nag forever
        Assert.Contains("UNKNOWN", SyncEventMapper.ToHistoryLine(e));
    }

    [Fact]
    public void Unknown_storage_event_carrying_an_error_code_is_a_full_failure()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.BrandNewSyncStep","Data.Error_Code":3758096686,"Data.Error_Description":"jerrcFileNodeFileCorrupt"}""");
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Null(issue.ExpiresUtc);   // a confirmed failure never self-clears; only a real success clears it
        Assert.Equal("0xE000012E", issue.Code);
        Assert.Null(issue.ExpiresUtc);             // a confirmed failure clears only on real success
    }

    [Theory]
    // real events this machine emits that the previous build dropped entirely
    [InlineData("""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated"}""")]
    [InlineData("""{"EventName":"Office.OneNote.UserInfoService.GetUserTypesRequestFailed"}""")]
    // plausible future ones
    [InlineData("""{"EventName":"Office.OneNote.Storage.UploadBlockedByPolicy"}""")]
    [InlineData("""{"EventName":"Office.OneNote.Sync.ContentStuckDetected"}""")]
    public void Failure_words_in_the_event_name_are_caught_even_with_no_fields(string json)
    {
        var e = Parse(json);
        Assert.Equal(SyncOutcome.SuspectedFailure, e.Outcome);
        Assert.True(e.IsProblem);
        var issue = SyncEventMapper.ToIssue(e, "etw")!;
                Assert.Contains("indicates a sync problem", issue.Summary);
    }

    [Fact]
    public void Malformed_storage_payload_is_reported_not_swallowed()
    {
        var e = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":tru""", out var malformed);
        Assert.True(malformed);
        Assert.NotNull(e);
        Assert.Equal(SyncOutcome.Unknown, e!.Outcome);
        Assert.Contains("could not be parsed", e.ClassificationReason);
    }

    // ---------- alternative field spellings ----------

    [Theory]
    [InlineData("""{"EventName":"Office.OneNote.Storage.X","Data.IsSuccess":false}""")]
    [InlineData("""{"EventName":"Office.OneNote.Storage.X","Data.Succeeded":false}""")]
    [InlineData("""{"EventName":"Office.OneNote.Storage.X","Data.WasSuccessful":false}""")]
    public void Alternative_success_field_spellings_are_honoured(string json) =>
        Assert.Equal(SyncOutcome.Failure, Parse(json).Outcome);

    [Theory]
    [InlineData("Data.SH_ErrorCode")]
    [InlineData("Data.Error_Code")]
    [InlineData("Data.NotebookErrorCode")]
    [InlineData("Data.SomePrefixErrorCode")]
    public void Any_error_code_field_spelling_is_detected(string field)
    {
        var e = Parse($$"""{"EventName":"Office.OneNote.Storage.X","{{field}}":3758096477}""");
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        Assert.Equal(0xE000005Du, e.ErrorCode);
    }

    [Fact]
    public void Zero_error_codes_and_no_error_texts_do_not_create_false_failures()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Success":true,"Data.NotebookErrorCode":0,"Data.Error":"No error","Data.OperationWithError":""}""");
        Assert.Equal(SyncOutcome.Success, e.Outcome);
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
    }

    // ---------- absent success flag ----------

    [Fact]
    public void Outcome_event_without_a_success_flag_is_Unknown_not_OK()
    {
        // the old build defaulted this to "OK"
        var e = Parse("""{"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Gosid":"{G}{1}"}""");
        Assert.Equal(SyncOutcome.Unknown, e.Outcome);
        Assert.True(e.IsProblem);
        Assert.False(e.IsSyncActivity);
    }

    // ---------- scope: benign non-Storage noise stays out ----------

    [Theory]
    [InlineData("""{"EventName":"Office.OneNote.System.AppLifeCycle.AppLaunch"}""")]
    [InlineData("""{"EventName":"Office.OneNote.Augmentation.Copilot.DisabledNoLicense"}""")]
    [InlineData("""{"EventName":"Office.OneNote.System.ConfigServiceReady"}""")]
    [InlineData("""{"EventName":"Office.Licensing.Something","Data.Error_Code":5}""")]
    [InlineData("""{"EventName":"Office.OneNote.Navigation.Navigate"}""")]
    public void Out_of_scope_events_are_ignored(string json) =>
        Assert.Null(SyncEventJson.TryParse("SendEvent " + json));

    [Fact]
    public void But_a_non_storage_onenote_event_with_a_real_error_is_picked_up()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Navigation.Navigate","Data.SH_ErrorCode":3758096477}""");
        Assert.Equal(SyncEventKind.OtherOneNoteSignal, e.Kind);
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
    }

    // ---------- transient handling ----------

    [Fact]
    public void A_transient_failure_is_still_an_error_by_default()
    {
        // two states: OneNote calling it "retryable" does not make it invisible
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":false,"Data.IsErrorTransient":true,"Data.Error_Code":3825336384,"Data.SectionResourceId_ResourceId":"x!s1"}""");
        Assert.Equal(SyncOutcome.Transient, e.Outcome);
        var issue = SyncEventMapper.ToIssue(e, "etw");
        Assert.NotNull(issue);
        Assert.Contains("has not recovered", issue!.Summary);
        Assert.Contains("FAILED(transient)", SyncEventMapper.ToHistoryLine(e));
    }

    [Fact]
    public void TransientPolicy_holds_back_only_codes_you_opted_into_until_the_threshold()
    {
        var t0 = DateTimeOffset.UtcNow;

        // a code NOT listed in [transient] is never held back — it is an error immediately
        Assert.False(TransientPolicy.Default.ShouldHoldBack("section:a", "0xE4020040", t0));

        var p = new TransientPolicy(["0xE402*"], threshold: 3, windowMinutes: 30);
        Assert.True(p.ShouldHoldBack("section:a", "0xE4020040", t0));                 // 1st: held back
        Assert.True(p.ShouldHoldBack("section:a", "0xE4020040", t0.AddMinutes(1)));   // 2nd: held back
        Assert.False(p.ShouldHoldBack("section:a", "0xE4020040", t0.AddMinutes(2)));  // 3rd: raise it
        // a different code is unaffected
        Assert.False(p.ShouldHoldBack("section:a", "0xE000005D", t0.AddMinutes(2)));

        p.Forget("section:a");
        Assert.True(p.ShouldHoldBack("section:a", "0xE4020040", t0.AddMinutes(3)));   // counter reset

        // occurrences outside the window never accumulate
        var q = new TransientPolicy(["0xE402*"], 3, 30);
        Assert.True(q.ShouldHoldBack("s", "0xE4020040", t0));
        Assert.True(q.ShouldHoldBack("s", "0xE4020040", t0.AddHours(1)));
        Assert.True(q.ShouldHoldBack("s", "0xE4020040", t0.AddHours(2)));
    }

    // ---------- the documented exception ----------

    [Fact]
    public void SyncScore_stays_diagnostic_and_is_the_only_error_bearing_exception()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SyncScore","Data.Error_Code":3758097184,"Data.Error_Description":"ErrCrypto_BadPassphrase"}""");
        Assert.Equal(SyncOutcome.Diagnostic, e.Outcome);
        Assert.False(e.IsProblem);
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
        Assert.Contains("SCORE", SyncEventMapper.ToHistoryLine(e));
        Assert.Contains("background replication scan", e.ClassificationReason);
    }

    // ---------- regression guard against the real event universe ----------

    [Theory]
    // every Office.OneNote.* event name observed on the target machine, with its expected outcome
    [InlineData("Office.OneNote.Storage.NotebookSyncResult", true)]
    [InlineData("Office.OneNote.Storage.SectionSyncResult", true)]
    [InlineData("Office.OneNote.Storage.SyncScore", true)]
    [InlineData("Office.OneNote.Storage.PageSyncSession", true)]
    [InlineData("Office.OneNote.Storage.ConnectivityChanged", true)]
    [InlineData("Office.OneNote.Storage.RealTime.NoteItHttpUpload", true)]
    [InlineData("Office.OneNote.Storage.RealTime.NoteItHttpDownload", true)]
    [InlineData("Office.OneNote.Storage.RealTime.NoteItService", true)]
    [InlineData("Office.OneNote.Storage.RealTime.SyncBlockerInstantiated", true)]
    [InlineData("Office.OneNote.UserInfoService.GetUserTypesRequestFailed", true)]
    [InlineData("Office.OneNote.Storage.AnythingNewMicrosoftAdds", true)]
    [InlineData("Office.OneNote.System.AppLifeCycle.AppLaunch", false)]
    [InlineData("Office.OneNote.Augmentation.Copilot.DisabledNoLicense", false)]
    public void Every_known_event_name_is_either_in_scope_or_deliberately_ignored(string name, bool inScope)
    {
        var e = SyncEventJson.TryParse($$"""SendEvent {"EventName":"{{name}}"}""");
        Assert.Equal(inScope, e is not null);
        if (e is not null)
        {
            // in scope means: classified, and never silently "successful" without evidence
            Assert.True(e.Outcome != SyncOutcome.Success || e.Kind is SyncEventKind.PageUpload or SyncEventKind.PageDownload);
            Assert.False(string.IsNullOrWhiteSpace(e.ClassificationReason));
        }
    }

    [Fact]
    public void Every_classified_problem_produces_an_actionable_issue()
    {
        string[] problems =
        [
            """{"EventName":"Office.OneNote.Storage.Whatever"}""",
            """{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated"}""",
            """{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":false}""",
            """{"EventName":"Office.OneNote.Storage.X","Data.Error_Code":3758096477}""",
            """{"EventName":"Office.OneNote.Storage.PageSyncSession","Data.ErrorState_TimeInActiveSyncErrorState":5000}""",
        ];
        foreach (var json in problems)
        {
            var e = Parse(json);
            Assert.True(e.IsProblem, e.EventName);
            var issue = SyncEventMapper.ToIssue(e, "etw");
            Assert.NotNull(issue);
            Assert.False(string.IsNullOrWhiteSpace(issue!.Summary), json);
            Assert.False(string.IsNullOrWhiteSpace(issue.Recommendation), json);
            Assert.Contains("event=", issue.TechnicalDetail);
        }
    }

    [Fact]
    public void An_unclassified_event_can_be_silenced_by_name_once_judged_benign()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated"}""");
        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Equal("Office.OneNote.Storage.RealTime.SyncBlockerInstantiated", issue.EventName);

        Assert.False(IgnoreRules.Empty.IsIgnored(issue));
        var rules = new IgnoreRules([], [], [], [], [], ["Office.OneNote.Storage.RealTime.SyncBlocker*"]);
        Assert.True(rules.IsIgnored(issue));
    }
}
