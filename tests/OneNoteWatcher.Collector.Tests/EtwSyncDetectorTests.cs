using OneNoteWatcher.Collector;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher.Collector.Tests;

public class EtwSyncDetectorTests
{
    private static OfficeLogMessage Msg(string sendEventJson) =>
        new(DateTime.UtcNow, OfficeEtw.TelemetryCategory, "SendEvent " + sendEventJson);

    private const string SectionFail =
        "{\"EventName\":\"Office.OneNote.Storage.SectionSyncResult\",\"Time\":\"2026-09-04T15:54:48.427Z\"," +
        "\"Data.Success\":false,\"Data.Error_Code\":3758096477,\"Data.Error_Description\":\"ErrFilePendingRename\"," +
        "\"Data.IsErrorTransient\":false,\"Data.SectionResourceId_ResourceId\":\"36B934175DC7E3A4!se6abe06d75b4\"}";
    private const string SectionOk =
        "{\"EventName\":\"Office.OneNote.Storage.SectionSyncResult\",\"Time\":\"2026-09-04T15:55:02.166Z\"," +
        "\"Data.Success\":true,\"Data.SectionResourceId_ResourceId\":\"36B934175DC7E3A4!se6abe06d75b4\"}";
    private const string NotebookOk =
        "{\"EventName\":\"Office.OneNote.Storage.NotebookSyncResult\",\"Time\":\"2026-09-04T15:55:04Z\"," +
        "\"Data.Success\":true,\"Data.NotebookErrorCode\":0,\"Data.Gosid\":\"{42AC7CD5-3122-4501-B9DF-254D9AC6D8F6}{1}\"}";
    private const string Offline = "{\"EventName\":\"Office.OneNote.Storage.ConnectivityChanged\",\"Data.InternetConnectivityNowAvailable\":false}";
    private const string Online  = "{\"EventName\":\"Office.OneNote.Storage.ConnectivityChanged\",\"Data.InternetConnectivityNowAvailable\":true}";

    private static (EtwSyncDetector det, string historyPath, List<WatcherStatus> statuses) NewDetector()
    {
        var dir = Path.Combine(Path.GetTempPath(), "onwatch_det_" + Guid.NewGuid().ToString("N"));
        var history = new SyncHistoryLog(dir);
        var statuses = new List<WatcherStatus>();
        // resolver pointed at empty dirs → names fall back to ids (deterministic in tests)
        var names = new NameResolver(new SearchIndexReader(Path.Combine(dir, "noidx")), new MruReader(Path.Combine(dir, "nomru")));
        var det = new EtwSyncDetector(new NullSource(), history, names,
            onStatusChanged: statuses.Add, oneNoteRunning: () => true);
        return (det, history.CurrentFile, statuses);
    }

    [Fact]
    public void Failure_then_recovery_raises_then_clears_issue_with_full_explanation()
    {
        var (det, historyPath, statuses) = NewDetector();

        det.OnMessage(Msg(SectionFail));
        var issue = Assert.Single(det.ActiveIssues());
        Assert.Equal("0xE000005D", issue.Code);
        Assert.Equal(Core.Diagnosis.FailureCategory.FileState, issue.Category);
        Assert.Contains("rename", issue.Summary);
        Assert.False(string.IsNullOrEmpty(issue.Recommendation));
        Assert.Contains("section-rid=", issue.TechnicalDetail);

        det.OnMessage(Msg(SectionOk));
        Assert.Empty(det.ActiveIssues());

        var history = File.ReadAllText(historyPath);
        Assert.Contains("FAILED  0xE000005D ErrFilePendingRename  [FileState]", history);
        Assert.Equal(new[] { 1, 0 }, statuses.Select(s => s.ActiveIssues.Count).ToArray());
    }

    [Fact]
    public void Repeated_failure_increments_occurrences_and_keeps_first_seen()
    {
        var (det, _, _) = NewDetector();
        det.OnMessage(Msg(SectionFail));
        det.OnMessage(Msg(SectionFail.Replace("15:54:48.427Z", "15:56:00.000Z")));
        var issue = Assert.Single(det.ActiveIssues());
        Assert.Equal(2, issue.Occurrences);
        Assert.Equal(new DateTimeOffset(2026, 9, 4, 15, 54, 48, 427, TimeSpan.Zero), issue.FirstSeen);
    }

    [Fact]
    public void Notebook_result_updates_status_and_offline_is_a_warning_that_clears()
    {
        var (det, _, statuses) = NewDetector();
        det.OnMessage(Msg(NotebookOk));
        var nb = Assert.Single(det.Snapshot().Notebooks);
        Assert.True(nb.LastSuccess);
        Assert.NotNull(det.Snapshot().LastSyncEventUtc);

        det.OnMessage(Msg(Offline));
        var off = Assert.Single(det.ActiveIssues());
        Assert.Equal(IssueKind.Offline, off.Kind);
                Assert.False(det.Snapshot().InternetAvailable);

        det.OnMessage(Msg(Online));
        Assert.Empty(det.ActiveIssues());
        Assert.True(det.Snapshot().InternetAvailable);
        Assert.True(statuses.Count >= 3);
    }

    [Fact]
    public void Page_upload_counts_as_notebook_activity_so_edits_move_last_sync()
    {
        // the user's edit reached the phone via the real-time channel while NotebookSyncResult stayed old
        var (det, historyPath, _) = NewDetector();
        det.OnMessage(Msg(NotebookOk)); // 15:55:04 on notebook {42AC7CD5}
        det.OnMessage(Msg("{\"EventName\":\"Office.OneNote.Storage.RealTime.NoteItHttpUpload\",\"Time\":\"2026-09-04T16:30:00Z\"," +
                          "\"Data.UploadTimeInMs\":900,\"Data.SectionId_ResourceId\":\"36B934175DC7E3A4!s1\",\"Data.NotebookId_ResourceId\":\"36B934175DC7E3A4!626\"}"));
        var s = det.Snapshot();
        Assert.Equal(new DateTimeOffset(2026, 9, 4, 16, 30, 0, TimeSpan.Zero), s.LastSyncEventUtc);
        Assert.Contains(s.Notebooks, n => n.LastSyncUtc == s.LastSyncEventUtc && n.LastSuccess == true);
        Assert.Contains("PAGE-UPLOAD  OK  (900 ms)", File.ReadAllText(historyPath));
    }

    [Fact]
    public void Non_sync_messages_are_ignored_cheaply()
    {
        var (det, _, _) = NewDetector();
        det.OnMessage(new OfficeLogMessage(DateTime.UtcNow, OfficeEtw.TelemetryCategory,
            "SendEvent {\"EventName\":\"Office.OneNote.Navigation.Navigate\"}"));
        Assert.Equal(1, det.MessagesSeen);
        Assert.Equal(0, det.SyncEventsParsed);
    }

    [Fact]
    public void Diag_log_backfill_ingests_finished_session_once()
    {
        var (det, historyPath, _) = NewDetector();
        var shared = Path.GetDirectoryName(historyPath)!;
        var diag = Path.Combine(shared, "diag"); Directory.CreateDirectory(diag);
        var row = "09/04/2026 22:54:48.427\tONENOTE (0x1)\t0x2\tMicrosoft OneNote\tTelemetry Event\tb7vzq\tMedium\tSendEvent " + SectionFail + "\t\r\n";
        var padded = "Timestamp\tProcess\n" + row + new string('\0', 500);
        File.WriteAllText(Path.Combine(diag, "Primary1_A.log"), padded);

        var bf = new DiagLogBackfill(shared, diag);
        Assert.Equal(1, bf.Run(det));
        Assert.Single(det.ActiveIssues());
        Assert.Equal(0, bf.Run(det)); // remembered → not re-ingested
    }


    [Fact]
    public void Unknown_event_becomes_a_visible_issue_and_only_a_proven_success_clears_its_own_scope()
    {
        var (det, historyPath, _) = NewDetector();
        det.OnMessage(Msg("""{"EventName":"Office.OneNote.Storage.SomeFutureThing"}"""));
        var issue = Assert.Single(det.ActiveIssues());
                Assert.NotNull(issue.ExpiresUtc);

        // a success on a DIFFERENT scope must not clear it
        det.OnMessage(Msg(NotebookOk));
        Assert.Single(det.ActiveIssues());
        Assert.Contains("UNKNOWN", File.ReadAllText(historyPath));
    }

    [Fact]
    public void Suspected_issue_is_shown_while_fresh_and_expires_once_stale()
    {
        var (det, _, _) = NewDetector();
        var now = DateTimeOffset.UtcNow.ToString("O");
        det.OnMessage(Msg($$"""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated","Time":"{{now}}"}"""));
        var issue = Assert.Single(det.Snapshot().ActiveIssues);   // visible while fresh
                Assert.NotNull(issue.ExpiresUtc);

        // the same signal from an old session is already past its TTL and is swept away,
        // so an unconfirmed suspicion can never pulse forever
        var (old, _, _) = NewDetector();
        old.OnMessage(Msg("""{"EventName":"Office.OneNote.Storage.RealTime.SyncBlockerInstantiated","Time":"2020-01-01T00:00:00Z"}"""));
        Assert.Empty(old.Snapshot().ActiveIssues);
    }

    [Fact]
    public void Confirmed_failure_does_not_expire_and_clears_only_on_success()
    {
        var (det, _, _) = NewDetector();
        det.OnMessage(Msg(SectionFail.Replace("2026-09-04T15:54:48.427Z", "2020-01-01T00:00:00Z")));
        Assert.Single(det.Snapshot().ActiveIssues);  // no TTL, survives the snapshot sweep
        det.OnMessage(Msg(SectionOk));
        Assert.Empty(det.ActiveIssues());
    }

    [Fact]
    public void A_configured_transient_code_is_held_back_until_its_threshold_then_raised()
    {
        var dir = Path.Combine(Path.GetTempPath(), "onwatch_tr_" + Guid.NewGuid().ToString("N"));
        var det = new EtwSyncDetector(new NullSource(), new SyncHistoryLog(dir),
            new NameResolver(new SearchIndexReader(Path.Combine(dir, "noidx")), new MruReader(Path.Combine(dir, "nomru"))),
            oneNoteRunning: () => true,
            transient: new Core.Rules.TransientPolicy(["0xE402*"], threshold: 3, windowMinutes: 30));
        var tr = """{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":false,"Data.IsErrorTransient":true,"Data.Error_Code":3825336384,"Data.SectionResourceId_ResourceId":"x!s1"}""";

        det.OnMessage(Msg(tr)); Assert.Empty(det.ActiveIssues());   // held back, still in history
        det.OnMessage(Msg(tr)); Assert.Empty(det.ActiveIssues());
        det.OnMessage(Msg(tr)); Assert.Single(det.ActiveIssues());  // threshold reached → error

        Assert.Contains("FAILED(transient)", File.ReadAllText(new SyncHistoryLog(dir).CurrentFile));
        try { Directory.Delete(dir, true); } catch { }
    }

    [Fact]
    public void An_unlisted_transient_code_is_an_error_immediately()
    {
        var (det, _, _) = NewDetector();   // default policy holds nothing back
        det.OnMessage(Msg("""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":false,"Data.IsErrorTransient":true,"Data.Error_Code":3825336384,"Data.SectionResourceId_ResourceId":"x!s9"}"""));
        Assert.Single(det.ActiveIssues());
    }

    [Fact]
    public void Dropped_etw_records_are_surfaced_not_silent()
    {
        var (det, _, _) = NewDetector();
        det.ReportEventsLost(42);
        var i = Assert.Single(det.ActiveIssues());
        Assert.Equal(IssueKind.SourceUnavailable, i.Kind);
        Assert.Contains("dropped", i.Summary);
        Assert.Equal(42, det.Snapshot().EventsLost);
    }

    [Fact]
    public void Malformed_payload_is_counted()
    {
        var (det, _, _) = NewDetector();
        det.OnMessage(new OfficeLogMessage(DateTime.UtcNow, OfficeEtw.TelemetryCategory,
            """SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.Success":tru"""));
        Assert.Equal(1, det.MalformedPayloads);
        Assert.Single(det.ActiveIssues());
    }

    private sealed class NullSource : IEtwMessageSource
    {
        public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct) { }
    }
}
