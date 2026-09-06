using OneNoteWatcher.Core;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Logging;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

public class RealtimeEventTests
{
    // real payloads observed in the diagnostic log on 2026-09-05
    private const string Upload = """SendEvent {"EventName": "Office.OneNote.Storage.RealTime.NoteItHttpUpload", "Time": "2026-09-05T15:17:54.485Z", "Data.UploadTimeInMs": 1024, "Data.SectionId_ResourceId": "36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213", "Data.NotebookId_ResourceId": "36B934175DC7E3A4!626"}""";
    private const string Download = """SendEvent {"EventName": "Office.OneNote.Storage.RealTime.NoteItHttpDownload", "Time": "2026-09-05T15:03:03.585Z", "Data.TimeToConfirmSyncedWithServerInMs": 2000, "Data.SectionId_ResourceId": "36B934175DC7E3A4!s7720fe23b57c4a15bb4940724cd05fe7", "Data.NotebookId_ResourceId": "36B934175DC7E3A4!626"}""";
    private const string ServiceOk = """SendEvent {"EventName": "Office.OneNote.Storage.RealTime.NoteItService", "Time": "2026-09-05T15:03:12.810Z", "Data.OperationWithError": "", "Data.Error": "No error", "Data.SectionId_ResourceId": "36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213"}""";
    private const string ServiceErr = """SendEvent {"EventName": "Office.OneNote.Storage.RealTime.NoteItService", "Time": "2026-09-05T15:03:12.810Z", "Data.OperationWithError": "Upload", "Data.Error": "HTTP 503 Service Unavailable", "Data.SectionId_ResourceId": "36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213"}""";

    [Fact]
    public void Page_upload_is_sync_activity_with_section_and_notebook_ids()
    {
        Assert.True(SyncEventJson.LooksLikeSyncSendEvent(Upload));
        var e = SyncEventJson.TryParse(Upload)!;
        Assert.Equal(SyncEventKind.PageUpload, e.Kind);
        Assert.True(e.IsSyncActivity);
        Assert.Equal(SyncOutcome.Success, e.Outcome);
        Assert.Equal(1024, e.TransferTimeMs);
        Assert.Equal("36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213", e.SectionResourceId);
        Assert.Equal("36B934175DC7E3A4!626", e.NotebookResourceId);
        Assert.Contains("PAGE-UPLOAD  OK  (1024 ms)", SyncEventMapper.ToHistoryLine(e, new ResolvedNames("Note", "Quick Notes", null)));
    }

    [Fact]
    public void Page_download_is_sync_activity()
    {
        var e = SyncEventJson.TryParse(Download)!;
        Assert.Equal(SyncEventKind.PageDownload, e.Kind);
        Assert.True(e.IsSyncActivity);
        Assert.Equal(2000, e.TransferTimeMs);
    }

    [Fact]
    public void Realtime_service_error_is_a_warning_issue_and_no_error_is_healthy()
    {
        var ok = SyncEventJson.TryParse(ServiceOk)!;
        Assert.Equal(SyncOutcome.Success, ok.Outcome); Assert.Null(SyncEventMapper.ToIssue(ok, "etw"));

        var bad = SyncEventJson.TryParse(ServiceErr)!;
        Assert.Equal(SyncOutcome.Failure, bad.Outcome);
        Assert.Equal("Upload: HTTP 503 Service Unavailable", bad.ErrorDescription);
        var issue = SyncEventMapper.ToIssue(bad, "etw", new ResolvedNames("Note", "Quick Notes", null))!;
                Assert.Contains("real-time", issue.Summary);
        Assert.Contains("REALTIME  FAILED", SyncEventMapper.ToHistoryLine(bad));
    }
}

public class AppLogTests
{
    [Fact]
    public void Writes_daily_file_and_purges_old_ones()
    {
        var dir = Path.Combine(Path.GetTempPath(), "onwatch_log_" + Guid.NewGuid().ToString("N"));
        try
        {
            var log = new AppLog(dir, "tray");
            log.Info("hello"); log.Warn("careful"); log.Error("boom", new InvalidOperationException("x"));
            var text = File.ReadAllText(log.CurrentFile);
            Assert.Contains("INFO  hello", text);
            Assert.Contains("ERROR boom: InvalidOperationException: x", text);

            var old = Path.Combine(dir, "tray-2020-01-01.log");
            File.WriteAllText(old, "old"); File.SetLastWriteTimeUtc(old, DateTime.UtcNow.AddDays(-30));
            var keep = Path.Combine(dir, "section-names.json"); File.WriteAllText(keep, "{}"); File.SetLastWriteTimeUtc(keep, DateTime.UtcNow.AddDays(-30));

            Assert.Equal(1, AppLog.Purge(dir, 14));
            Assert.False(File.Exists(old));
            Assert.True(File.Exists(keep));            // only *.log / *.txt are purged
            Assert.True(File.Exists(log.CurrentFile)); // today's file kept
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }

    [Fact]
    public void History_uses_daily_files()
    {
        var dir = Path.Combine(Path.GetTempPath(), "onwatch_hist_" + Guid.NewGuid().ToString("N"));
        try
        {
            var h = new SyncHistoryLog(dir);
            h.Append("line");
            Assert.EndsWith($"sync-history-{DateTime.Now:yyyy-MM-dd}.log", h.CurrentFile);
            Assert.Contains("line", File.ReadAllText(h.CurrentFile));
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }
}
