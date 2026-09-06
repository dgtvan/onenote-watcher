using System.Diagnostics;
using OneNoteWatcher.Core.Config;
using OneNoteWatcher.Core.Health;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Logging;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Rules;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher;

/// <summary>
/// The unelevated tray app. The icon ALWAYS reflects the last known state — whether or not OneNote is
/// running — from the collector's status.json plus local detectors (OAlerts dialogs, Graph outcome
/// baseline, health checks). TWO STATES ONLY:
///   red pulsing = ERROR — something is wrong and you need to look (where/why/fix in the balloon and window)
///   green       = SUCCESS — everything the watcher can check is healthy
/// There is no middle "warning" tier, because that is the tier people learn to ignore. Noise is
/// controlled explicitly instead, through [ignore] rules, the [transient] policy, and TTLs.
/// Menu is deliberately minimal: Show issues &amp; status / Check now / Open config.ini / Quit.
/// </summary>
public sealed class TrayApp : IDisposable
{
    private readonly NotifyIcon _tray = new();
    private readonly System.Windows.Forms.Timer _pollTimer = new();
    private readonly System.Windows.Forms.Timer _pulseTimer = new();
    private readonly System.Windows.Forms.Timer _graphTimer = new();
    private readonly OAlertsDetector _oalerts;
    private readonly GraphAuth? _graphAuth;
    private readonly GraphPoller? _graph;
    private readonly AppLog _log;

    private readonly string _sharedDir, _logsDir, _statusPath, _configPath;
    private readonly IgnoreRules _ignore;
    private readonly bool _simulate;

    private bool _pulseBright;
    private TrayIcons.State _state = TrayIcons.State.Error;   // fail-closed until proven healthy
    private string _stateReason = "";
    private Icon? _current;
    private WatcherStatus? _status;
    private readonly Dictionary<string, SyncIssue> _localIssues = new();
    private IReadOnlyList<SyncIssue> _graphIssues = [];
    /// <summary>Dedupe keys already reported as cloud-cleared, so the log says it once.</summary>
    private readonly HashSet<string> _cloudCleared = new(StringComparer.Ordinal);
    private readonly SyncHistoryLog _history;
    private List<SyncIssue> _issues = [];
    private string _lastBalloonKey = "";
    private bool _graphBusy;
    private IssuesForm? _issuesForm;
    private string? _configError;
    private readonly bool _graphEnabled;
    private readonly bool _etwEnabled;
    private bool _oneNoteWasRunning;

    public TrayApp(bool simulate)
    {
        _simulate = simulate;
        _configPath = Path.Combine(AppContext.BaseDirectory, "config.ini");
        IniFile ini;
        try { ini = File.Exists(_configPath) ? IniFile.Load(_configPath) : IniFile.Parse(""); }
        catch (Exception ex) { ini = IniFile.Parse(""); _configError = ex.Message; }
        if (!File.Exists(_configPath)) _configError ??= "file not found";
        _ignore = IgnoreRules.FromConfig(ini);

        _sharedDir = ini.Get("etw", "shared_dir", AppContext.BaseDirectory.TrimEnd('\\')); // one root: exes, config, status, logs
        _logsDir = Path.Combine(_sharedDir, "logs");
        _statusPath = Path.Combine(_sharedDir, "status.json");
        try { Directory.CreateDirectory(_logsDir); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        _log = new AppLog(_logsDir, "tray");
        _history = new SyncHistoryLog(_logsDir);
        _log.Info($"===== tray start  pid={Environment.ProcessId} user={Environment.UserName} simulate={simulate} config={_configPath}");
        AppLog.Purge(_logsDir, ini.GetInt("general", "log_retention_days", 14));

        var clientId = ini.Get("graph", "client_id");
        _etwEnabled = ini.GetBool("etw", "enabled", true);
        // keep the null test inline so the compiler can see clientId is non-null below
        if (ini.GetBool("graph", "enabled", true) && !string.IsNullOrWhiteSpace(clientId))
        {
            _graphEnabled = true;
            _graphAuth = new GraphAuth(clientId, ini.Get("graph", "tenant", "consumers"));
            _graph = new GraphPoller(_graphAuth, TimeSpan.FromMinutes(ini.GetInt("general", "grace_minutes", 10)), _sharedDir, _log);
            _graphTimer.Interval = Math.Max(1, ini.GetInt("general", "poll_minutes", 5)) * 60_000;
            _graphTimer.Tick += async (_, _) => { try { await PollGraphAsync(); } catch (Exception ex) { _log.Error("graph timer", ex); } };
            _graphTimer.Start();
        }
        else _log.Info("Graph disabled or no client_id");

        _oalerts = new OAlertsDetector(issue =>
        {
            _log.Warn($"OAlerts dialog: {issue.Message}");
            lock (_localIssues) _localIssues[issue.DedupeKey] = issue;
            try { _tray.ContextMenuStrip?.BeginInvoke(Refresh); } catch { }
        });
        _log.Info($"OAlerts watcher: {(_oalerts.Available ? "on" : "unavailable")}");

        _tray.Text = "OneNote Sync Watcher";
        _tray.Visible = true;
        _tray.ContextMenuStrip = BuildMenu();
        _tray.DoubleClick += (_, _) => ShowIssues();

        _pollTimer.Interval = 3000;
        _pollTimer.Tick += (_, _) => SafeRefresh();
        _pollTimer.Start();
        _pulseTimer.Interval = 500;
        _pulseTimer.Tick += (_, _) => { _pulseBright = !_pulseBright; ApplyIcon(); };

        Refresh();
        if (_graph is not null) _ = PollGraphAsync();
    }

    private ContextMenuStrip BuildMenu()
    {
        var m = new ContextMenuStrip();
        m.Items.Add("Show issues && status", null, (_, _) => ShowIssues());
        m.Items.Add("Check now", null, async (_, _) => { Refresh(); await PollGraphAsync(); });
        m.Items.Add(new ToolStripSeparator());
        m.Items.Add("Open config.ini", null, (_, _) => OpenPath(_configPath));
        m.Items.Add(new ToolStripSeparator());
        m.Items.Add("Quit", null, (_, _) => { _log.Info("quit from menu"); _tray.Visible = false; Application.Exit(); });
        return m;
    }

    private async Task SignInAsync()
    {
        if (_graphAuth is null) return;
        try
        {
            _log.Info("Graph sign-in started");
            await _graphAuth.SignInAsync((url, code) =>
            {
                _tray.ContextMenuStrip?.BeginInvoke(() =>
                {
                    try { Clipboard.SetText(code); } catch { }
                    MessageBox.Show(
                        $"1. Open {url}\n2. Enter this code (already copied to clipboard):\n\n      {code}\n\n" +
                        "3. Sign in with your personal Microsoft account and accept 'Read your OneNote notebooks'.\n\n" +
                        "This dialog can be closed; the watcher completes sign-in in the background.",
                        "OneNote Sync Watcher — sign in to Microsoft Graph", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                });
            }, CancellationToken.None);
            _log.Info("Graph sign-in succeeded");
            ShowBalloon("Signed in to Microsoft Graph", "Section names and the cloud-side check are now active.");
            await PollGraphAsync();
        }
        catch (Exception ex)
        {
            _log.Error("Graph sign-in failed", ex);
            MessageBox.Show("Sign-in failed: " + ex.Message, "OneNote Sync Watcher");
        }
    }

    private async Task SignOutAsync()
    {
        if (_graphAuth is null) return;
        await _graphAuth.SignOutAsync();
        _log.Info("Graph signed out");
        await PollGraphAsync();
    }


    private async Task PollGraphAsync()
    {
        if (_graph is null || _graphBusy || _simulate) return;
        _graphBusy = true;
        // the collector's own results go in, so a section OneNote has already reconciled is not
        // reported as a stranded change on the strength of two lagging timestamps
        var sections = _status is not null && DateTimeOffset.UtcNow - _status.UpdatedUtc < TimeSpan.FromMinutes(2)
            ? _status.Sections : null;
        try { _graphIssues = await _graph.PollAsync(OneNoteRunning(), _status?.InternetAvailable, CancellationToken.None, sections); }
        catch (Exception ex) { _log.Error("graph poll failed", ex); _graphIssues = []; }
        finally { _graphBusy = false; }
        Refresh();
    }


    /// <summary>
    /// Drop a collector error that the cloud check has since DISPROVED.
    ///
    /// OneNote does not reliably emit a success event after it recovers from a real-time upload error
    /// (observed: ErrOutOfSyncWithStore, then no further Storage telemetry for that section at all), so an
    /// error raised from telemetry alone can stand forever even though the content is safely in OneDrive.
    /// The cloud check is an independent source that can prove it: if OneDrive holds content for that
    /// section stamped later than the failure, the failure is over.
    ///
    /// This is still fail-closed — it removes an error only on positive proof from the server. No Graph,
    /// no stale poll, no section name, or no server movement all mean the error stays.
    /// </summary>
    private bool KeepAfterCloudCheck(SyncIssue issue)
    {
        if (_graph is null || issue.SectionName is null) return true;
        if (issue.Kind is IssueKind.UploadStuck or IssueKind.DownloadStuck) return true;  // the cloud check owns these
        if (!_graph.CloudConfirmedAfter(issue.NotebookName, issue.SectionName, issue.LastSeen)) return true;

        if (_cloudCleared.Add(issue.DedupeKey))
        {
            _log.Info($"cleared by cloud check: {issue.Location} — OneDrive holds content newer than the "
                    + $"failure at {issue.LastSeen.ToLocalTime():HH:mm:ss}, so the change did reach the server");
            // the history must not end on "FAILED" for a section that recovered
            _history.Append($"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}  {issue.Location}  RECOVERED  "
                          + $"cloud check confirms OneDrive has this section's content (failure at "
                          + $"{issue.LastSeen.ToLocalTime():HH:mm:ss} is resolved)");
        }
        return false;
    }

    /// <summary>
    /// Tray-side fail-closed checks. Anything that stops the watcher from SEEING becomes a visible
    /// issue — the app must never look healthy while it is blind.
    /// </summary>
    private IEnumerable<SyncIssue> HealthChecks()
    {
        if (_simulate) yield break;
        var now = DateTimeOffset.Now;

        if (_configError is not null)
            yield return HealthIssues.ConfigUnreadable(_configPath, _configError, now);

        if (!_etwEnabled)
            yield return HealthIssues.DetectorDisabled("Real-time sync monitoring (etw)", now);
        else if (_status is null || !_status.CollectorRunning || DateTimeOffset.UtcNow - _status.UpdatedUtc > TimeSpan.FromMinutes(2))
            yield return HealthIssues.CollectorDown(_status?.UpdatedUtc, now);

        if (!_graphEnabled)
            yield return HealthIssues.DetectorDisabled("Cloud-side check (graph)", now);
        else if (_graph is not null)
        {
            if (!_graph.SignedIn)
                yield return _graphAuth?.State == GraphAuthState.Expired
                    ? HealthIssues.GraphTokenExpired(now)      // had access and lost it → ALERT
                    : HealthIssues.GraphBlind("not signed in", now);
            else if (_graph.LastError is not null && _graph.LastFetchUtc is null)
                yield return HealthIssues.GraphBlind($"every attempt failed: {_graph.LastError}", now);
            else if (_graph.LastFetchUtc is { } f && DateTimeOffset.UtcNow - f > TimeSpan.FromMinutes(90))
                yield return HealthIssues.GraphBlind($"last successful check {f.ToLocalTime():HH:mm}", now);
        }

        if (!_oalerts.Available)
            yield return HealthIssues.OAlertsUnavailable(now);

        if (_log.WriteFailures > 0)
            yield return HealthIssues.LogWriteFailing(_log.WriteFailures, now);
    }

    private void SafeRefresh()
    {
        try { Refresh(); }
        catch (Exception ex)
        {
            // a broken refresh must be loud, not a frozen icon
            _log.Error("refresh failed", ex);
            _state = TrayIcons.State.Error;
            _tray.Text = "OneNote Sync Watcher — internal error, see tray log";
            ApplyIcon();
        }
    }

    private void Refresh()
    {
        _status = _simulate ? null : WatcherStatus.Load(_statusPath);

        var list = new List<SyncIssue>();
        if (_simulate) list.Add(SimulatedIssue());
        if (_status is not null) list.AddRange(_status.ActiveIssues.Where(KeepAfterCloudCheck));
        lock (_localIssues) list.AddRange(_localIssues.Values);
        list.AddRange(HealthChecks());
        var collectorFresh = _status is not null && DateTimeOffset.UtcNow - _status.UpdatedUtc < TimeSpan.FromMinutes(2);
        // NOTE: a notebook-level success from the collector is deliberately NOT used to suppress a
        // section-level "stranded change" finding — OneNote reports section errors separately inside a
        // notebook result, so notebook success does not prove the section uploaded. The cloud check
        // clears it on direct evidence instead (server timestamp catches up).
        list.AddRange(_graphIssues);
        _issues = list;

        // TWO STATES: any active, un-ignored issue means ERROR. Nothing to report means SUCCESS.
        var errors = _issues.Where(i => !_ignore.IsIgnored(i)).ToList();
        var collectorAlive = _status is not null && _status.CollectorRunning && collectorFresh;
        var oneNote = OneNoteRunning();

        // OneNote emits no final sync on shutdown (verified), so the moment it closes we re-check the
        // cloud side: anything edited but not uploaded is stranded until it is reopened.
        if (_oneNoteWasRunning && !oneNote)
        {
            _log.Info("OneNote closed — running an immediate cloud check for unsynced changes");
            _ = PollGraphAsync();
        }
        _oneNoteWasRunning = oneNote;

        TrayIcons.State newState;
        string detail;
        if (errors.Count > 0)
        {
            newState = TrayIcons.State.Error;
            var top = errors[0];
            detail = errors.Count == 1
                ? $"{(top.Location.Length > 0 ? top.Location + " — " : "")}{Describe(top)}"
                : $"{errors.Count} problems: {(top.Location.Length > 0 ? top.Location + " — " : "")}{Describe(top)}";
        }
        else
        {
            newState = TrayIcons.State.Ok;
            detail = $"in sync — last activity {LastActivityLocal()}"
                     + (oneNote ? "" : " (OneNote not running)")
                     + (collectorAlive ? "" : " (collector off; cloud check only)");
        }

        // the signature includes the issue count, so a change in WHAT is wrong is logged even when the
        // headline message and colour stay the same
        var signature = $"{newState}|{detail}|n{errors.Count}";
        if (signature != _stateReason)
        {
            var names = string.Join(" | ", errors.Take(3).Select(i =>
                $"{i.Detector}:{(i.Location.Length > 0 ? i.Location + " " : "")}{i.Message}"));
            _log.Info($"icon {(newState != _state ? $"{_state} → {newState}" : $"{newState} (unchanged)")}: {detail}" +
                      $"  [errors={errors.Count}{(names.Length > 0 ? " :: " + names : "")}]");
            _state = newState; _stateReason = signature;
            if (_state == TrayIcons.State.Error) _pulseTimer.Start(); else { _pulseTimer.Stop(); _pulseBright = false; }
        }

        if (_state == TrayIcons.State.Error && errors.Count > 0)
        {
            var a = errors[0];
            if (a.DedupeKey != _lastBalloonKey)
            {
                _lastBalloonKey = a.DedupeKey;
                ShowBalloon($"OneNote sync problem — {(a.Location.Length > 0 ? a.Location : a.Detector)}",
                    $"{a.Code} {a.Message}\n{a.Summary}\nFix: {a.Recommendation}");
            }
        }
        else if (_state != TrayIcons.State.Error) _lastBalloonKey = "";

        var tip = "OneNote Sync Watcher — " + detail;
        _tray.Text = tip.Length > 127 ? tip[..124] + "…" : tip;
        ApplyIcon();
    }

    /// <summary>Best available one-line description of an issue for the tooltip.</summary>
    private static string Describe(SyncIssue i) =>
        !string.IsNullOrEmpty(i.Summary) ? i.Summary
        : !string.IsNullOrEmpty(i.Message) ? i.Message
        : i.Kind.ToString();

    private bool AllLastOk() => _status?.Notebooks.All(n => n.LastSuccess != false) ?? true;
    private string LastActivityLocal() => _status?.LastSyncEventUtc is { } t ? t.ToLocalTime().ToString("HH:mm") : "unknown";

    private void ApplyIcon()
    {
        var next = TrayIcons.Make(_state, _pulseBright);
        _tray.Icon = next; _current?.Dispose(); _current = next;
    }

    private static bool OneNoteRunning() => Process.GetProcessesByName("ONENOTE").Length > 0;

    private void ShowIssues()
    {
        if (_issuesForm is { IsDisposed: false })
        {
            _issuesForm.Render();
            if (_issuesForm.WindowState == FormWindowState.Minimized) _issuesForm.WindowState = FormWindowState.Normal;
            _issuesForm.Activate(); return;
        }
        _issuesForm = new IssuesForm(
            () => new IssuesView(_issues, _status, _ignore.IsIgnored, _graph),
            SignInAsync, SignOutAsync, () => OpenPath(_logsDir));
        _issuesForm.Show();
    }

    private static void OpenPath(string path)
    {
        try { Process.Start(new ProcessStartInfo(path) { UseShellExecute = true }); }
        catch (Exception ex) { MessageBox.Show($"Couldn't open {path}\n{ex.Message}", "OneNote Sync Watcher"); }
    }

    private void ShowBalloon(string title, string text)
    {
        _tray.BalloonTipTitle = title.Length > 63 ? title[..63] : title;
        _tray.BalloonTipText = text.Length > 255 ? text[..252] + "…" : text;
        _tray.ShowBalloonTip(8000);
    }

    private static SyncIssue SimulatedIssue() => new()
    {
        Detector = "simulate", Kind = IssueKind.ErrorCodeReported,
        NotebookName = "Watcher Test", SectionName = "Demo Section",
        Code = "0xE000005D", Message = "ErrFilePendingRename (simulated)",
        Category = Core.Diagnosis.FailureCategory.FileState,
        Summary = "The section file is pending a rename/move on OneDrive.",
        Recommendation = "Usually clears on the next sync; if it persists, close and reopen the notebook.",
        FirstSeen = DateTimeOffset.Now, LastSeen = DateTimeOffset.Now,
    };

    public void Dispose()
    {
        _log.Info("tray dispose");
        _pollTimer.Dispose(); _pulseTimer.Dispose(); _graphTimer.Dispose(); _oalerts.Dispose();
        _tray.Visible = false; _tray.Dispose(); _current?.Dispose();
    }
}
