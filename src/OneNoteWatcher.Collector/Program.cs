using OneNoteWatcher.Collector;
using OneNoteWatcher.Core.Config;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Logging;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

// OneNoteWatcher.Collector — the elevated real-time ETW consumer.
//   (no args)           live mode: real-time session on OfficeLoggingLiblet (needs admin)
//   --replay <etl>      offline: replay a captured .etl (no admin) — verification/tests
//   --no-backfill       skip the diagnostic-log backfill of finished sessions
//   --shared-dir <d>    status/section map/logs root (default: [etw] shared_dir in config.ini, else the exe's folder)
//   --user-profile <d>  the OneNote user's profile (e.g. C:\Users\vandang); the collector runs as SYSTEM,
//                       so %LOCALAPPDATA% would be SYSTEM's. install.ps1 passes this.
//   --config <file>     config.ini to read log_retention_days from (default: next to the exe)

string? replay = null; bool backfill = true; string? userProfile = null; string? sharedArg = null;
string configPath = Path.Combine(AppContext.BaseDirectory, "config.ini");
for (var i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--replay" when i + 1 < args.Length: replay = args[++i]; break;
        case "--shared-dir" when i + 1 < args.Length: sharedArg = args[++i]; break;
        case "--user-profile" when i + 1 < args.Length: userProfile = args[++i]; break;
        case "--config" when i + 1 < args.Length: configPath = args[++i]; break;
        case "--no-backfill": backfill = false; break;
    }
}
IniFile ini; string? configError = null;
try { ini = File.Exists(configPath) ? IniFile.Load(configPath) : IniFile.Parse(""); }
catch (Exception ex) { ini = IniFile.Parse(""); configError = ex.Message; }
if (!File.Exists(configPath)) configError ??= "file not found";
// everything (exes, config, status, logs) lives in ONE root: the install folder
var sharedDir = sharedArg ?? ini.Get("etw", "shared_dir", AppContext.BaseDirectory.TrimEnd('\\'));
var retentionDays = ini.GetInt("general", "log_retention_days", 14);
var logsDir = Path.Combine(sharedDir, "logs");
var localAppData = userProfile is not null
    ? Path.Combine(userProfile, "AppData", "Local")
    : Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

if (args.Contains("--audit"))
{
    // classification coverage report over the real diagnostic logs — no ETW, no admin needed
    Console.WriteLine(ClassificationAudit.Run(Path.Combine(localAppData, @"Temp\Diagnostics\ONENOTE")));
    return 0;
}

Directory.CreateDirectory(logsDir);
AppLog.EchoToConsole = Environment.UserInteractive && !Console.IsOutputRedirected;
var log = new AppLog(logsDir, "collector");
var history = new SyncHistoryLog(logsDir);
var statusPath = Path.Combine(sharedDir, "status.json");
log.Info($"===== collector start  pid={Environment.ProcessId} user={Environment.UserName} mode={(replay is null ? "live" : "replay " + replay)}");
log.Info($"sharedDir={sharedDir}  logs={logsDir}  retention={retentionDays}d  config={configPath}");
log.Info($"user LocalAppData={localAppData}  index={Directory.Exists(Path.Combine(localAppData, @"Microsoft\OneNote\16.0\FullTextSearchIndex"))}  mru={Directory.Exists(Path.Combine(localAppData, @"Microsoft\Office\16.0\MruServiceCache"))}  diaglogs={Directory.Exists(Path.Combine(localAppData, @"Temp\Diagnostics\ONENOTE"))}");
log.Info($"purged {AppLog.Purge(logsDir, retentionDays)} old log file(s)");

var liveSource = replay is null ? new LiveEtwMessageSource() : null;
IEtwMessageSource source = (IEtwMessageSource?)liveSource ?? new EtlFileMessageSource(replay!);

var sectionMapPath = Path.Combine(sharedDir, "section-names.json");
var sectionMap = OneNoteWatcher.Core.Graph.SectionNameMap.Load(sectionMapPath);
var sectionMapLoaded = DateTime.UtcNow;
log.Info($"section-name map: {sectionMap.Count} keys ({(sectionMap.Count == 0 ? "sign in to Graph from the tray to get section names" : "ok")})");
var names = new NameResolver(
    new SearchIndexReader(Path.Combine(localAppData, @"Microsoft\OneNote\16.0\FullTextSearchIndex")),
    new MruReader(Path.Combine(localAppData, @"Microsoft\Office\16.0\MruServiceCache")))
{
    SectionNameByResourceId = (rid, gosid) =>
    {
        if (DateTime.UtcNow - sectionMapLoaded > TimeSpan.FromMinutes(2))
        {
            var reloaded = OneNoteWatcher.Core.Graph.SectionNameMap.Load(sectionMapPath);
            if (reloaded.Count != sectionMap.Count) log.Info($"section-name map reloaded: {reloaded.Count} keys");
            sectionMap = reloaded; sectionMapLoaded = DateTime.UtcNow;
        }
        return sectionMap.Lookup(rid, gosid);
    }
};
names.Refresh(TimeSpan.Zero);

var detector = new EtwSyncDetector(source, history, names, log: log,
    onStatusChanged: status => status.Save(statusPath),
    transient: OneNoteWatcher.Core.Rules.TransientPolicy.FromConfig(ini));
if (liveSource is not null) liveSource.OnEventsLost = detector.ReportEventsLost;

var backfiller = new DiagLogBackfill(sharedDir, Path.Combine(localAppData, @"Temp\Diagnostics\ONENOTE"));
if (backfill)
{
    var n = backfiller.Run(detector, log);
    log.Info($"backfilled {n} finished session log(s); events so far={detector.SyncEventsParsed}");
}

// config problems are surfaced, never silently defaulted away (ignore rules would not be applied)
if (configError is not null)
{
    log.Warn($"config.ini unreadable ({configError}) — running on defaults");
    detector.AddHealthIssue(OneNoteWatcher.Core.Health.HealthIssues.KeyConfigUnreadable,
        OneNoteWatcher.Core.Health.HealthIssues.ConfigUnreadable(configPath, configError, DateTimeOffset.UtcNow));
}

// SILENCE IS NOT SUCCESS: if OneNote is running and no sync telemetry arrives within this window,
// either OneNote stopped syncing or our pipeline is broken — both raise an alert.
var activityTimeout = TimeSpan.FromMinutes(
    ini.GetInt("general", "pipeline_timeout_minutes", ini.GetInt("general", "activity_timeout_minutes", 45)));
log.Info($"pipeline watchdog: alert if NO Office telemetry at all for {activityTimeout.TotalMinutes:F0} min while OneNote runs " +
         "(measured max gap on a healthy machine: 17 min)");

// heartbeat + health sweep + periodic summary
var lastSummaryEvents = 0;
using var heartbeat = new Timer(_ =>
{
    try
    {
        // OneNote emits no final sync on shutdown, so when it closes we re-ingest the session log that
        // just unlocked — anything the live ETW session missed is still captured.
        if (backfill && detector.OneNoteJustClosed())
        {
            log.Info("OneNote closed — re-ingesting the finished session log");
            var n = backfiller.Run(detector, log);
            log.Info($"post-exit backfill: {n} session log(s), events total={detector.SyncEventsParsed}");
        }
        detector.EvaluateHealth(activityTimeout);
        detector.Snapshot().Save(statusPath);
        if (DateTime.Now.Minute % 10 == 0 && detector.SyncEventsParsed != lastSummaryEvents)
        {
            lastSummaryEvents = detector.SyncEventsParsed;
            log.Info($"summary: messages={detector.MessagesSeen} syncEvents={detector.SyncEventsParsed} activeIssues={detector.ActiveIssues().Count} onenote={detector.Snapshot().OneNoteRunning} eventsLost={detector.EventsLost} malformed={detector.MalformedPayloads}");
        }
        if (DateTime.Now.Hour == 3 && DateTime.Now.Minute == 0) AppLog.Purge(logsDir, retentionDays);
    }
    catch (Exception ex) { log.Error("heartbeat failed", ex); }
}, null, TimeSpan.FromSeconds(5), TimeSpan.FromSeconds(30));

using var cts = new CancellationTokenSource();
Console.CancelKeyPress += (_, e) => { e.Cancel = true; cts.Cancel(); };
AppDomain.CurrentDomain.ProcessExit += (_, _) => log.Info("process exit");
// a crash must leave evidence and mark the collector as down, never look healthy
AppDomain.CurrentDomain.UnhandledException += (_, e) =>
{
    log.Error("UNHANDLED EXCEPTION — collector is stopping", e.ExceptionObject as Exception);
    try { (detector.Snapshot() with { CollectorRunning = false }).Save(statusPath); } catch { }
};

// live mode: keep the session alive across transient failures
var attempt = 0;
while (!cts.IsCancellationRequested)
{
    try
    {
        log.Info(replay is null ? "starting ETW session OfficeLoggingLiblet" : "replaying");
        detector.Run(cts.Token);
        break; // replay finished or session stopped cleanly
    }
    catch (Exception ex)
    {
        log.Error("ETW session failed" + (attempt == 0 ? " (need admin?)" : ""), ex);
        (detector.Snapshot() with { CollectorRunning = false }).Save(statusPath);
        if (replay is not null || ++attempt > 20) return 2;
        Thread.Sleep(TimeSpan.FromSeconds(Math.Min(60, 5 * attempt)));
    }
}

(detector.Snapshot() with { CollectorRunning = replay is not null }).Save(statusPath);
log.Info($"done. messages={detector.MessagesSeen} syncEvents={detector.SyncEventsParsed} activeIssues={detector.ActiveIssues().Count}");
return 0;
