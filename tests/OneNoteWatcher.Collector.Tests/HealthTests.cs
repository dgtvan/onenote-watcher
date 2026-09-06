using OneNoteWatcher.Collector;
using OneNoteWatcher.Core.Health;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Collector.Tests;

/// <summary>
/// System-wide fail-closed rules: the watcher must never look healthy while it cannot actually see.
/// "Silence is not success."
/// </summary>
public class HealthTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "onwatch_health_" + Guid.NewGuid().ToString("N"));

    /// <summary>Simulated awake time, advanced by tests in step with their simulated wall clock.</summary>
    private TimeSpan _awake = TimeSpan.Zero;

    private EtwSyncDetector NewDetector(bool oneNoteRunning, string? indexDir = null)
    {
        var names = new NameResolver(
            new SearchIndexReader(indexDir ?? Path.Combine(_dir, "missing-index")),
            new MruReader(Path.Combine(_dir, "nomru")));
        names.Refresh(TimeSpan.Zero);
        return new EtwSyncDetector(new NullSource(), new SyncHistoryLog(_dir), names,
            oneNoteRunning: () => oneNoteRunning, awakeClock: () => _awake);
    }

    private static OfficeLogMessage Msg(string json) =>
        new(DateTime.UtcNow, OfficeEtw.TelemetryCategory, "SendEvent " + json);

    private const string NotebookOk =
        """{"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Success":true,"Data.Gosid":"{G}{1}"}""";

    [Fact]
    public void No_telemetry_at_all_while_OneNote_runs_is_an_error()
    {
        var det = NewDetector(oneNoteRunning: true);
        det.OnMessage(Msg(NotebookOk));                       // telemetry flowing "now"

        det.EvaluateHealth(TimeSpan.FromMinutes(45));         // fresh → nothing wrong
        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));

        // pretend the pipeline went silent for hours WHILE THE MACHINE WAS AWAKE
        _awake += TimeSpan.FromHours(3);
        det.EvaluateHealth(TimeSpan.FromMinutes(45), DateTimeOffset.UtcNow.AddHours(3));
        var dead = Assert.Single(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
        Assert.Contains("pipeline", dead.Summary);
        Assert.Contains("Collector", dead.Recommendation);
    }

    [Fact]
    public void An_idle_OneNote_that_simply_is_not_syncing_is_NOT_an_error()
    {
        // measured: idle OneNote goes up to 114 min between SYNC events while telemetry keeps flowing.
        // Watching sync events would fire every time OneNote is left open and idle.
        var det = NewDetector(oneNoteRunning: true);
        var quiet = new OfficeLogMessage(DateTime.UtcNow, OfficeEtw.TelemetryCategory,
            """SendEvent {"EventName":"Office.OneNote.Navigation.Navigate"}""");   // telemetry, but no sync
        det.OnMessage(quiet);

        det.EvaluateHealth(TimeSpan.FromMinutes(45));
        // (the fixture points at a missing index dir, so only assert about the liveness check)
        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
        Assert.Null(det.Snapshot().LastSyncEventUtc);          // no sync at all — still fine
    }

    [Fact]
    public void A_restart_never_reports_silence_it_did_not_actually_watch()
    {
        // the real false alarm: after a restart the backfill sets last-sync from HISTORICAL data, and
        // the watchdog claimed 2.6 h of silence 5 seconds after starting
        var det = NewDetector(oneNoteRunning: true);
        det.Ingest(OneNoteWatcher.Core.Parsing.SyncEventJson.TryParse(
            """SendEvent {"EventName":"Office.OneNote.Storage.NotebookSyncResult","Time":"2020-01-01T00:00:00Z","Data.Success":true,"Data.Gosid":"{G}{1}"}""")!);

        det.EvaluateHealth(TimeSpan.FromMinutes(45));          // uptime is seconds, not years
        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
    }

    [Fact]
    public void Silence_while_OneNote_is_closed_is_expected_and_not_an_error()
    {
        var det = NewDetector(oneNoteRunning: false);
        det.EvaluateHealth(TimeSpan.FromMinutes(45), DateTimeOffset.UtcNow.AddHours(5));
        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
    }

    [Fact]
    public void Watchdog_clears_once_telemetry_resumes()
    {
        var det = NewDetector(oneNoteRunning: true);
        _awake += TimeSpan.FromMinutes(30);
        det.EvaluateHealth(TimeSpan.FromMinutes(1), DateTimeOffset.UtcNow.AddMinutes(30));
        Assert.Contains(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));

        det.OnMessage(Msg(NotebookOk));
        det.EvaluateHealth(TimeSpan.FromMinutes(45));
        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
    }

    [Fact]
    public void Time_the_machine_spent_ASLEEP_is_not_reported_as_silence()
    {
        // 2026-09-06: a 3 h sleep raised "no Office telemetry at all for 3.0 h". Nothing was wrong —
        // the watcher simply was not running to observe those hours.
        var det = NewDetector(oneNoteRunning: true);
        det.OnMessage(Msg(NotebookOk));

        // wall clock jumps three hours; awake time does not advance at all
        det.EvaluateHealth(TimeSpan.FromMinutes(45), DateTimeOffset.UtcNow.AddHours(3));

        Assert.DoesNotContain(det.ActiveIssues(), i => i.Message.Contains("no telemetry"));
    }

    [Fact]
    public void An_unreadable_index_is_surfaced_not_ignored()
    {
        var det = NewDetector(oneNoteRunning: true);          // index dir does not exist
        det.EvaluateHealth(TimeSpan.FromHours(99));
        var issue = Assert.Single(det.ActiveIssues(), i => i.Message.Contains("local index"));
                Assert.Contains("incomplete", issue.Summary);
    }

    [Fact]
    public void Dropped_etw_records_surface_and_use_the_shared_health_key()
    {
        var det = NewDetector(oneNoteRunning: true);
        det.ReportEventsLost(7);
        var i = Assert.Single(det.ActiveIssues(), x => x.Message.Contains("dropped"));
        Assert.Contains(HealthIssues.KeyEventsLost, i.TechnicalDetail);
        Assert.Equal(7, det.Snapshot().EventsLost);
    }

    [Fact]
    public void Externally_detected_problems_can_be_published()
    {
        var det = NewDetector(oneNoteRunning: true);
        det.AddHealthIssue(HealthIssues.KeyConfigUnreadable,
            HealthIssues.ConfigUnreadable(@"C:\x\config.ini", "file not found", DateTimeOffset.Now));
        var i = Assert.Single(det.ActiveIssues(), x => x.Message.Contains("config.ini"));
        Assert.Contains("ignore rules are NOT applied", i.Summary);
    }

    [Fact]
    public void Health_issues_all_carry_an_explanation_and_a_fix()
    {
        var now = DateTimeOffset.Now;
        SyncIssue[] all =
        [
            HealthIssues.PipelineSilent(TimeSpan.FromHours(2), TimeSpan.FromHours(1), now),
            HealthIssues.CollectorDown(now, now),
            HealthIssues.IndexUnavailable("locked", now),
            HealthIssues.GraphBlind("not signed in", now),
            HealthIssues.OAlertsUnavailable(now),
            HealthIssues.LogWriteFailing(3, now),
            HealthIssues.DetectorDisabled("Cloud-side check (graph)", now),
            HealthIssues.ConfigUnreadable("c:\\x", "denied", now),
            HealthIssues.EventsLost(5, now),
        ];
        Assert.All(all, i =>
        {
            Assert.False(string.IsNullOrWhiteSpace(i.Summary));
            Assert.False(string.IsNullOrWhiteSpace(i.Recommendation));
                    });
    }

    private sealed class NullSource : IEtwMessageSource
    {
        public void Process(Action<OfficeLogMessage> onMessage, CancellationToken ct) { }
    }

    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }

    [Fact]
    public void Dead_pipeline_alert_survives_the_real_status_file_round_trip_to_the_tray()
    {
        // integration: the collector's health sweep -> Snapshot() -> status.json on disk -> what the tray
        // actually loads. This is the seam the end-to-end test exercises with the real binaries.
        var det = NewDetector(oneNoteRunning: true);
        var statusPath = Path.Combine(_dir, "status.json");

        _awake += TimeSpan.FromHours(3);   // three hours of RUNNING time, not three hours of sleep
        det.EvaluateHealth(TimeSpan.FromMinutes(60), DateTimeOffset.UtcNow.AddHours(3));
        det.Snapshot().Save(statusPath);

        var asTheTraySeesIt = Core.Status.WatcherStatus.Load(statusPath);
        Assert.NotNull(asTheTraySeesIt);
        var alert = Assert.Single(asTheTraySeesIt!.ActiveIssues, i => i.Message.Contains("no telemetry"));
        Assert.Contains("no telemetry", alert.Message);
        Assert.Contains("pipeline", alert.Summary);
        Assert.False(string.IsNullOrWhiteSpace(alert.Recommendation));
    }

    [Fact]
    public void An_expired_microsoft_sign_in_is_an_alert_not_a_warning()
    {
        // the cloud check is the ONLY thing that can prove edits reached OneDrive after OneNote closes,
        // so losing it silently is the worst case
        var expired = HealthIssues.GraphTokenExpired(DateTimeOffset.Now);
                Assert.Contains("expired", expired.Message);
        Assert.Contains("would NOT be detected", expired.Summary);
        Assert.Contains("Sign in", expired.Recommendation);

        // never having signed in is setup, not a regression → warning
                // both use the same key so one replaces the other rather than stacking up
        Assert.Contains(HealthIssues.KeyGraphBlind, expired.TechnicalDetail);
        Assert.Contains(HealthIssues.KeyGraphBlind, HealthIssues.GraphBlind("x", DateTimeOffset.Now).TechnicalDetail);
    }
}
