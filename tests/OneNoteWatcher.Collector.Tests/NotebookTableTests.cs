using OneNoteWatcher.Collector;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;

namespace OneNoteWatcher.Collector.Tests;

/// <summary>
/// The NOTEBOOKS table is a summary of the same truth as the PROBLEMS list, so it must never show a
/// notebook as healthy (or omit it entirely) while an error against it is open. Reported 2026-09-06:
/// a real-time failure on Note / eSim Data left "Note" missing from the table altogether, because the
/// table was only ever written on success.
/// </summary>
public class NotebookTableTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "onwatch_nbtable_" + Guid.NewGuid().ToString("N"));

    private EtwSyncDetector NewDetector() =>
        new(new NullSource(), new SyncHistoryLog(_dir),
            new NameResolver(new SearchIndexReader(Path.Combine(_dir, "noidx")), new MruReader(Path.Combine(_dir, "nomru"))),
            oneNoteRunning: () => true);

    private static OfficeLogMessage Msg(string json) =>
        new(DateTime.UtcNow, OfficeEtw.TelemetryCategory, "SendEvent " + json);

    private const string Nb = "36B934175DC7E3A4!626";
    private const string Sec = "36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2";

    private static string RealTimeFail(string at) =>
        $$"""{"EventName":"Office.OneNote.Storage.RealTime.NoteItService","Time":"{{at}}","Data.OperationWithError":"Http Patch for Upload Failed","Data.Error":"Win32Error: ErrOutOfSyncWithStore (0xE000002E) tag_4oxx9","Data.SectionId_ResourceId":"{{Sec}}","Data.NotebookId_ResourceId":"{{Nb}}"}""";
    private static string SectionOk(string at) =>
        $$"""{"EventName":"Office.OneNote.Storage.SectionSyncResult","Time":"{{at}}","Data.Success":true,"Data.SectionResourceId_ResourceId":"{{Sec}}","Data.NotebookId_ResourceId":"{{Nb}}"}""";

    [Fact]
    public void A_notebook_whose_only_event_is_a_failure_still_gets_a_row()
    {
        var det = NewDetector();
        det.OnMessage(Msg(RealTimeFail("2026-09-06T08:33:19Z")));

        var row = Assert.Single(det.Snapshot().Notebooks);
        Assert.False(row.LastSuccess);
        Assert.Contains("ErrOutOfSyncWithStore", row.LastError);
    }

    [Fact]
    public void A_failure_flips_a_previously_healthy_notebook_to_failed()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionOk("2026-09-06T08:30:00Z")));
        Assert.True(Assert.Single(det.Snapshot().Notebooks).LastSuccess);

        det.OnMessage(Msg(RealTimeFail("2026-09-06T08:33:19Z")));

        var row = Assert.Single(det.Snapshot().Notebooks);
        Assert.False(row.LastSuccess);
    }

    [Fact]
    public void The_row_goes_back_to_ok_once_a_covering_success_clears_the_error()
    {
        var det = NewDetector();
        det.OnMessage(Msg(RealTimeFail("2026-09-06T08:33:19Z")));
        Assert.False(Assert.Single(det.Snapshot().Notebooks).LastSuccess);

        det.OnMessage(Msg(SectionOk("2026-09-06T08:35:00Z")));

        var row = Assert.Single(det.Snapshot().Notebooks);
        Assert.True(row.LastSuccess);
        Assert.Null(row.LastError);
        Assert.Empty(det.Snapshot().ActiveIssues);
    }

    [Fact]
    public void The_table_never_disagrees_with_the_problem_list()
    {
        var det = NewDetector();
        det.OnMessage(Msg(RealTimeFail("2026-09-06T08:33:19Z")));

        var snap = det.Snapshot();
        var failedNotebooks = snap.Notebooks.Where(n => n.LastSuccess == false).Select(n => n.Name).ToHashSet();
        var notebooksWithIssues = snap.ActiveIssues.Where(i => i.NotebookName is not null).Select(i => i.NotebookName!).ToHashSet();
        Assert.Equal(notebooksWithIssues, failedNotebooks);
    }

    // ---- per-section results published for the cloud check ----

    [Fact]
    public void A_completed_section_sync_is_published_under_its_canonical_key()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionOk("2026-09-06T08:28:45Z")));

        var pub = Assert.Single(det.Snapshot().Sections);
        Assert.Equal(OneNoteWatcher.Core.Model.SectionKey.Normalize(Sec), pub.Key);
        Assert.Equal(new DateTimeOffset(2026, 9, 6, 8, 28, 45, TimeSpan.Zero), pub.LastHealthySyncUtc);
    }

    [Fact]
    public void Real_time_and_page_activity_also_count_as_healthy_sync()
    {
        var det = NewDetector();
        // The index is re-stamped by ANY reconciliation, so any healthy section activity explains the
        // stamp. Measured 2026-09-06: a section went hours with real-time activity and no section sync,
        // so requiring a full section sync made the explanation unobtainable.
        det.OnMessage(Msg($$"""{"EventName":"Office.OneNote.Storage.RealTime.NoteItHttpUpload","Time":"2026-09-06T08:28:45Z","Data.UploadTimeInMs":900,"Data.SectionId_ResourceId":"{{Sec}}"}"""));
        det.OnMessage(Msg($$"""{"EventName":"Office.OneNote.Storage.RealTime.NoteItService","Time":"2026-09-06T08:28:46Z","Data.Error":"No error","Data.SectionId_ResourceId":"{{Sec}}"}"""));

        Assert.Equal(new DateTimeOffset(2026, 9, 6, 8, 28, 46, TimeSpan.Zero),
            Assert.Single(det.Snapshot().Sections).LastHealthySyncUtc);
    }

    [Fact]
    public void A_failed_event_never_counts_as_healthy_sync()
    {
        var det = NewDetector();
        det.OnMessage(Msg(RealTimeFail("2026-09-06T08:33:19Z")));

        Assert.Empty(det.Snapshot().Sections);
    }

    [Fact]
    public void Only_the_newest_section_sync_is_kept()
    {
        var det = NewDetector();
        det.OnMessage(Msg(SectionOk("2026-09-06T08:35:00Z")));
        det.OnMessage(Msg(SectionOk("2026-09-06T08:28:45Z")));   // replayed by the backfill, out of order

        Assert.Equal(new DateTimeOffset(2026, 9, 6, 8, 35, 0, TimeSpan.Zero),
            Assert.Single(det.Snapshot().Sections).LastHealthySyncUtc);
    }

    [Fact]
    public void Results_survive_a_restart_via_the_published_status()
    {
        var before = NewDetector();
        before.OnMessage(Msg(SectionOk("2026-09-06T08:28:45Z")));

        var after = NewDetector();
        Assert.Empty(after.Snapshot().Sections);
        after.SeedSectionSuccesses(before.Snapshot().Sections);

        Assert.Equal(new DateTimeOffset(2026, 9, 6, 8, 28, 45, TimeSpan.Zero),
            Assert.Single(after.Snapshot().Sections).LastHealthySyncUtc);
    }

    private sealed class NullSource : IEtwMessageSource
    {
        public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct) { }
    }

    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }
}
