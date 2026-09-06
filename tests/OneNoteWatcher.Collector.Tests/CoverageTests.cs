using OneNoteWatcher.Collector;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Collector.Tests;

/// <summary>
/// An error is cleared ONLY by a later success that genuinely covers it. A success that proves nothing
/// about the failed content must leave the error standing — clearing it would silently lose a real
/// sync failure, which is the one outcome this app exists to prevent.
/// </summary>
public class CoverageTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "onwatch_cov_" + Guid.NewGuid().ToString("N"));

    private EtwSyncDetector NewDetector() =>
        new(new NullSource(), new SyncHistoryLog(_dir),
            new NameResolver(new SearchIndexReader(Path.Combine(_dir, "noidx")), new MruReader(Path.Combine(_dir, "nomru"))),
            oneNoteRunning: () => true);

    private static OfficeLogMessage Msg(string json) =>
        new(DateTime.UtcNow, OfficeEtw.TelemetryCategory, "SendEvent " + json);

    private const string Sec = "36B934175DC7E3A4!sAAA";
    private const string OtherSec = "36B934175DC7E3A4!sBBB";

    private static string SectionFail(string at, string rid = Sec) =>
        $$"""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Time":"{{at}}","Data.Success":false,"Data.Error_Code":3758096477,"Data.Error_Description":"ErrFilePendingRename","Data.SectionResourceId_ResourceId":"{{rid}}"}""";
    private static string SectionOk(string at, string rid = Sec) =>
        $$"""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Time":"{{at}}","Data.Success":true,"Data.SectionResourceId_ResourceId":"{{rid}}"}""";
    private static string NotebookOk(string at) =>
        $$"""{"EventName":"Office.OneNote.Storage.NotebookSyncResult","Time":"{{at}}","Data.Success":true,"Data.NotebookErrorCode":0,"Data.Gosid":"{G}{1}"}""";
    private static string PageUploadOk(string at, string rid = Sec) =>
        $$"""{"EventName":"Office.OneNote.Storage.RealTime.NoteItHttpUpload","Time":"{{at}}","Data.UploadTimeInMs":900,"Data.SectionId_ResourceId":"{{rid}}"}""";
    private static string RealTimeOk(string at, string rid = Sec) =>
        $$"""{"EventName":"Office.OneNote.Storage.RealTime.NoteItService","Time":"{{at}}","Data.Error":"No error","Data.SectionId_ResourceId":"{{rid}}"}""";

    // ---------- what DOES cover a section failure ----------

    [Fact]
    public void A_later_section_sync_success_covers_that_sections_failure()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        Assert.Single(det.ActiveIssues());

        det.OnMessage(Msg(SectionOk("2026-09-06T10:05:00Z")));
        Assert.Empty(det.ActiveIssues());          // same section, later, full sync completed → covered
    }

    // ---------- what does NOT ----------

    [Fact]
    public void A_page_upload_does_NOT_cover_a_section_failure()
    {
        // the previous build cleared the whole section on one page upload — one page reaching the
        // server says nothing about the rest of the section
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        det.OnMessage(Msg(PageUploadOk("2026-09-06T10:05:00Z")));
        Assert.Single(det.ActiveIssues());
    }

    [Fact]
    public void A_notebook_success_does_NOT_cover_a_section_failure()
    {
        // OneNote's own NotebookSyncResult carries separate IsSectionErrorSuppressed /
        // IsSectionErrorUnexpected fields: a notebook sync can succeed while its sections had errors
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        det.OnMessage(Msg(NotebookOk("2026-09-06T10:05:00Z")));
        Assert.Single(det.ActiveIssues());
    }

    [Fact]
    public void A_success_on_a_DIFFERENT_section_does_not_cover_it()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        det.OnMessage(Msg(SectionOk("2026-09-06T10:05:00Z", OtherSec)));
        Assert.Single(det.ActiveIssues());
    }

    [Fact]
    public void A_real_time_success_does_not_cover_a_full_section_sync_failure()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        det.OnMessage(Msg(RealTimeOk("2026-09-06T10:05:00Z")));
        Assert.Single(det.ActiveIssues());        // different mechanism, weaker evidence
    }

    // ---------- time ordering ----------

    [Fact]
    public void An_EARLIER_success_never_clears_a_later_failure()
    {
        // the post-exit backfill replays finished session logs, so out-of-order arrival is real
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:10:00Z")));
        det.OnMessage(Msg(SectionOk("2026-09-06T10:00:00Z")));    // older success arrives afterwards
        Assert.Single(det.ActiveIssues());
    }

    [Fact]
    public void A_stale_failure_replayed_after_a_covering_success_does_not_reopen_the_error()
    {
        // exactly what happens when OneNote closes and the collector re-ingests the session log
        var det = NewDetector();
        det.OnMessage(Msg(SectionOk("2026-09-06T10:05:00Z")));
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));   // older failure, already covered
        Assert.Empty(det.ActiveIssues());
    }

    // ---------- recovery after a self-healing condition ----------

    [Fact]
    public void Offline_then_online_recovers_cleanly()
    {
        var det = NewDetector();
        det.OnMessage(Msg("""{"EventName":"Office.OneNote.Storage.ConnectivityChanged","Data.InternetConnectivityNowAvailable":false}"""));
        Assert.Single(det.ActiveIssues());
        det.OnMessage(Msg("""{"EventName":"Office.OneNote.Storage.ConnectivityChanged","Data.InternetConnectivityNowAvailable":true}"""));
        Assert.Empty(det.ActiveIssues());
    }

    [Fact]
    public void A_repeat_failure_keeps_the_original_first_seen_and_still_clears_on_real_coverage()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionFail("2026-09-06T10:00:00Z")));
        det.OnMessage(Msg(SectionFail("2026-09-06T10:02:00Z")));
        var issue = Assert.Single(det.ActiveIssues());
        Assert.Equal(2, issue.Occurrences);
        Assert.Equal(new DateTimeOffset(2026, 9, 6, 10, 0, 0, TimeSpan.Zero), issue.FirstSeen);

        det.OnMessage(Msg(SectionOk("2026-09-06T10:03:00Z")));
        Assert.Empty(det.ActiveIssues());
    }

    [Fact]
    public void An_unclassified_event_is_never_cleared_by_any_success_only_by_its_TTL()
    {
        // we cannot know what would prove an unknown event resolved, so no success may clear it
        var det = NewDetector();
        var now = DateTimeOffset.UtcNow.ToString("O");
        det.OnMessage(Msg($$"""{"EventName":"Office.OneNote.Storage.SomethingNew","Time":"{{now}}"}"""));
        Assert.Single(det.ActiveIssues());

        det.OnMessage(Msg(SectionOk(DateTimeOffset.UtcNow.AddMinutes(1).ToString("O"))));
        det.OnMessage(Msg(NotebookOk(DateTimeOffset.UtcNow.AddMinutes(1).ToString("O"))));
        var issue = Assert.Single(det.ActiveIssues());
        Assert.NotNull(issue.ExpiresUtc);          // clears only by TTL
    }

    private sealed class NullSource : IEtwMessageSource
    {
        public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct) { }
    }

    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }
}
