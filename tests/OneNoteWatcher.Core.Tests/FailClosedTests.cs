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
        Assert.Equal("0xE000012E", issue.Code);
    }

    [Fact]
    public void A_confirmed_failure_clears_only_on_success_unless_no_success_could_ever_reach_it()
    {
        // A located failure is held until proof: SyncScope keys it by section, and a section success
        // covers it. No TTL — the error waits for evidence.
        var located = Parse("""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":false,"Data.Error_Code":3758096477,"Data.SectionResourceId_ResourceId":"x!s1"}""");
        Assert.Null(SyncEventMapper.ToIssue(located, "etw")!.ExpiresUtc);

        // An UNRECOGNISED event is keyed by event NAME, and CoveredBy never returns an event key — so
        // no success in the system can ever clear it. Without a TTL that is not "held pending proof",
        // it is stuck red forever. Time is its only exit, so it gets one.
        var unclearable = Parse("""{"EventName":"Office.OneNote.Storage.BrandNewSyncStep","Data.Error_Code":3758096686}""");
        Assert.Equal(SyncOutcome.Failure, unclearable.Outcome);
        Assert.NotNull(SyncEventMapper.ToIssue(unclearable, "etw")!.ExpiresUtc);

        // …and it is not a way to go quiet on a problem that is still happening: every recurrence
        // re-arms the TTL from the new event's time.
        var t0 = DateTimeOffset.UtcNow;
        var later = SyncEventMapper.ToIssue(unclearable with { Time = t0.AddHours(5) }, "etw")!;
        Assert.True(later.ExpiresUtc > t0 + SyncEventMapper.SuspectedIssueTtl);
    }

    [Theory]
    // real events this machine emits that the previous build dropped entirely
    [InlineData("""{"EventName":"Office.OneNote.Storage.RealTime.ContentSyncBlockerInstantiated"}""")]
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

    // ---------- known non-outcome telemetry: benign, but only while it reports nothing wrong ----------

    [Theory]
    [InlineData("Office.OneNote.Storage.RealTime.FileDataObjectDownload")]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoViaCobalt")]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoStats")]
    [InlineData("Office.OneNote.Storage.RealTime.FdoDownloadRequestBlockerInstantiated")]
    [InlineData("Office.OneNote.Storage.RealTime.DoesNotebookSatisfyNoteItPrerequisites")]
    public void Attachment_download_chatter_is_not_a_sync_failure(string name)
    {
        // Reported 2026-09-06 23:01: opening pages with attachments produced five unknown event names
        // and pinned the tray red for the 6 h TTL, while the same sections logged REALTIME OK in the
        // same second. Each carries no success flag and no error field — it reports no outcome at all.
        var e = Parse($$"""{"EventName":"{{name}}","Time":"2026-09-06T16:01:27Z"}""");
        Assert.Equal(SyncOutcome.Diagnostic, e.Outcome);
        Assert.False(e.IsProblem);
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
    }

    [Theory]
    [InlineData("Office.OneNote.Storage.RealTime.FileDataObjectDownload")]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoViaCobalt")]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoStats")]
    [InlineData("Office.OneNote.Storage.RealTime.FdoDownloadRequestBlockerInstantiated")]
    [InlineData("Office.OneNote.Storage.RealTime.DoesNotebookSatisfyNoteItPrerequisites")]
    public void But_the_same_event_reporting_a_real_error_is_still_a_failure(string name)
    {
        // the whole reason these live below the error checks and not in DiagnosticEvents: an attachment
        // that genuinely fails to download must still reach the user, with its code
        var e = Parse($$"""{"EventName":"{{name}}","Data.Error_Code":3758096477}""");
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
        Assert.True(e.IsProblem);
        Assert.Equal("0xE000005D", SyncEventMapper.ToIssue(e, "etw")!.Code);
    }

    [Fact]
    public void A_request_blocker_is_a_wait_handle_but_an_unlisted_blocker_is_still_suspect()
    {
        // "Blocker" in a name is a failure word, but OneNote also uses it for internal blocking-wait
        // handles. The exemption is deliberately narrow — only "RequestBlocker". An unrecognised
        // blocker event still surfaces; only names with evidence behind them are on the benign list.
        Assert.Equal(SyncOutcome.Unknown,
            Parse("""{"EventName":"Office.OneNote.Storage.RealTime.SomeFutureRequestBlockerInstantiated"}""").Outcome);
        Assert.Equal(SyncOutcome.SuspectedFailure,
            Parse("""{"EventName":"Office.OneNote.Storage.RealTime.ContentSyncBlockerInstantiated"}""").Outcome);
    }

    [Fact]
    public void The_startup_sync_gate_is_not_a_sync_failure()
    {
        // Observed once per OneNote launch on 2026-09-06 (×2) and 2026-09-08, one second after the
        // connectivity ONLINE events and 0-2 s before a successful PAGE-DOWNLOAD. The captured payload
        // carries no Data.* fields at all — it cannot name anything as blocked.
        // verbatim from the UNCLASSIFIED PAYLOAD log line of 2026-09-08 10:20:42 (token elided),
        // spaced separators and envelope fields included — the parser must cope with the real shape
        var e = Parse("""{"EventName": "Office.OneNote.Storage.RealTime.SyncBlockerInstantiated", "Flags": 30962273224818945, "InternalSequenceNumber": 260, "Time": "2026-09-08T03:20:40Z", "AriaTenantToken": "<elided>"}""");
        Assert.Equal(SyncOutcome.Diagnostic, e.Outcome);
        Assert.False(e.IsProblem);
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
    }

    [Fact]
    public void The_benign_list_outranks_the_failure_word_in_the_name_but_not_real_error_evidence()
    {
        // ordering proof: SyncBlockerInstantiated trips the name heuristic, so being on the list has to
        // win for it to go quiet — and error evidence has to win over the list.
        Assert.Equal(SyncOutcome.Diagnostic,
            Parse("""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated"}""").Outcome);
        Assert.Equal(SyncOutcome.Failure,
            Parse("""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated","Data.Error_Code":3758096477}""").Outcome);
    }

    // ---------- an HTTP status is an outcome, and must be read as one ----------

    [Fact]
    public void A_failing_http_status_is_read_as_the_outcome_instead_of_being_guessed_from_the_name()
    {
        // 2026-09-08: this arrived with Data.HttpStatus 503 and was reported as "indicates a problem,
        // but without a definite outcome" — the definite outcome was sitting unread in the payload.
        // verbatim from the UNCLASSIFIED PAYLOAD log line of 2026-09-08 10:20:47 (token elided)
        var e = Parse("""{"EventName": "Office.OneNote.UserInfoService.GetUserTypesRequestFailed", "Flags": 30962273224818945, "InternalSequenceNumber": 535, "Time": "2026-09-08T03:20:45Z", "AriaTenantToken": "<elided>", "Data.HttpStatus": 503}""");
        Assert.Equal(SyncOutcome.Transient, e.Outcome);        // 5xx is the server's own "try again"
        Assert.Equal("HTTP 503", e.ErrorDescription);

        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Equal(Core.Diagnosis.FailureCategory.Network, issue.Category);
        Assert.Contains("HttpStatus=503", issue.TechnicalDetail);
        Assert.NotNull(issue.ExpiresUtc);                      // event-scoped: nothing else could clear it
        Assert.Equal(new DateTimeOffset(2026, 9, 8, 3, 20, 45, TimeSpan.Zero) + SyncEventMapper.SuspectedIssueTtl,
            issue.ExpiresUtc);                                 // anchored to the event, not to parse time
    }

    [Theory]
    [InlineData(500, true)]
    [InlineData(503, true)]
    [InlineData(408, true)]
    [InlineData(429, true)]
    [InlineData(404, false)]
    [InlineData(403, false)]
    public void Server_side_and_throttling_statuses_are_transient_client_errors_are_not(int status, bool transient)
    {
        var e = Parse($$"""{"EventName":"Office.OneNote.Storage.SomeServiceCall","Data.HttpStatus":{{status}}}""");
        Assert.Equal(transient ? SyncOutcome.Transient : SyncOutcome.Failure, e.Outcome);
    }

    [Fact]
    public void A_successful_http_status_is_positive_evidence_not_just_the_absence_of_an_error()
    {
        // without this a 200 would fall through to the fail-closed default and be reported as a problem
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SomeServiceCall","Data.HttpStatus":200}""");
        Assert.Equal(SyncOutcome.Success, e.Outcome);
        Assert.False(e.IsProblem);
    }

    [Fact]
    public void A_real_error_code_still_outranks_a_healthy_http_status()
    {
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SomeServiceCall","Data.HttpStatus":200,"Data.Error_Code":3758096477}""");
        Assert.Equal(SyncOutcome.Failure, e.Outcome);
    }

    // ---------- an unclassified event carries the evidence needed to classify it ----------

    [Fact]
    public void An_unclassified_event_keeps_its_payload_so_the_report_can_be_acted_on()
    {
        // OneNote holds its diagnostic log exclusively locked while it runs, so a warning that only
        // names the event leaves nothing to classify it from once the moment has passed.
        var e = Parse("""{"EventName":"Office.OneNote.Storage.SomeFutureThing","Data.Whatever":42}""");
        Assert.Equal(SyncOutcome.Unknown, e.Outcome);
        Assert.Contains("\"Data.Whatever\":42", e.RawPayload);

        // a malformed payload is the case where the evidence matters most
        var bad = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":tru""")!;
        Assert.Contains("Data.Success", bad.RawPayload);
    }

    [Fact]
    public void A_classified_event_carries_no_payload()
    {
        // only the events that need reporting pay the memory cost
        Assert.Null(Parse("""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":true}""").RawPayload);
        Assert.Null(Parse("""{"EventName":"Office.OneNote.Storage.SyncScore"}""").RawPayload);
        Assert.Null(Parse("""{"EventName":"Office.OneNote.Storage.RealTime.DownloadFdoStats"}""").RawPayload);
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
    // observed 2026-09-06 when pages with attachments were opened
    [InlineData("Office.OneNote.Storage.RealTime.FileDataObjectDownload", true)]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoViaCobalt", true)]
    [InlineData("Office.OneNote.Storage.RealTime.DownloadFdoStats", true)]
    [InlineData("Office.OneNote.Storage.RealTime.FdoDownloadRequestBlockerInstantiated", true)]
    [InlineData("Office.OneNote.Storage.RealTime.DoesNotebookSatisfyNoteItPrerequisites", true)]
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
            """{"EventName":"Office.OneNote.Storage.RealTime.ContentSyncBlockerInstantiated"}""",
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
        var e = Parse("""{"EventName":"Office.OneNote.Storage.RealTime.ContentSyncBlockerInstantiated"}""");
        var issue = SyncEventMapper.ToIssue(e, "etw")!;
        Assert.Equal("Office.OneNote.Storage.RealTime.ContentSyncBlockerInstantiated", issue.EventName);

        Assert.False(IgnoreRules.Empty.IsIgnored(issue));
        var rules = new IgnoreRules([], [], [], [], [], ["Office.OneNote.Storage.RealTime.ContentSyncBlocker*"]);
        Assert.True(rules.IsIgnored(issue));
    }
}
