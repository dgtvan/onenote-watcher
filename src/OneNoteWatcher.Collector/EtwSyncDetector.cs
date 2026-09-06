using System.Runtime.CompilerServices;
using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Health;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Logging;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;
using OneNoteWatcher.Core.Rules;
using OneNoteWatcher.Core.Status;

[assembly: InternalsVisibleTo("OneNoteWatcher.Collector.Tests")]

namespace OneNoteWatcher.Collector;

/// <summary>
/// PRIMARY real-time detector, FAIL-CLOSED (docs/fail-closed.md): every OneNote event is
/// classified, only a proven success counts as success, and anything else becomes a visible error.
///
/// Clearing is as strict as raising: an error is removed ONLY when a later success genuinely COVERS
/// it (see <see cref="SyncScope"/>). A success that proves nothing about the failed content leaves
/// the error standing.
/// </summary>
public sealed class EtwSyncDetector
{
    public const string DetectorName = "etw";

    /// <summary>An open error plus the evidence needed to decide whether a later success covers it.</summary>
    private sealed record ActiveEntry(SyncIssue Issue, string ScopeKey, string? SectionRid, DateTimeOffset FailedAt);

    private readonly IEtwMessageSource _source;
    private readonly SyncHistoryLog _history;
    private readonly NameResolver _names;
    private readonly AppLog? _log;
    private readonly TransientPolicy _transient;
    private readonly Action<WatcherStatus>? _onStatusChanged;
    private readonly Func<bool> _oneNoteRunning;
    /// <summary>Awake-time source; injectable so the sleep rule can be tested against a simulated clock.</summary>
    private readonly Func<TimeSpan> _awakeNow;

    private readonly Dictionary<string, ActiveEntry> _active = new(StringComparer.Ordinal);
    /// <summary>Latest proven success per scope — stops a stale failure (e.g. replayed by the post-exit
    /// backfill) from re-opening an error that a newer success already covered.</summary>
    private readonly Dictionary<string, DateTimeOffset> _coveredUntil = new(StringComparer.Ordinal);
    private readonly Dictionary<string, NotebookStatus> _notebooks = new(StringComparer.OrdinalIgnoreCase);
    /// <summary>Last healthy sync activity per canonical section key — published for the cloud check,
    /// which cannot otherwise tell a stranded change from a section OneNote has already reconciled.</summary>
    private readonly Dictionary<string, SectionSyncState> _sectionSuccess = new(StringComparer.Ordinal);
    private readonly HashSet<string> _seenUnknownEvents = new(StringComparer.Ordinal);
    private readonly object _gate = new();

    private bool? _internet;
    private DateTimeOffset? _lastActivity;
    /// <summary>When the source last delivered ANY telemetry — the pipeline-liveness signal.</summary>
    private DateTimeOffset? _lastMessageAt;
    private readonly DateTimeOffset _startedUtc = DateTimeOffset.UtcNow;
    /// <summary>Awake-time stamps for the liveness watchdog, so a suspended machine is not counted as
    /// silence we observed (see <see cref="AwakeClock"/>).</summary>
    private readonly TimeSpan _startedAwake;
    private TimeSpan _lastMessageAwake;
    private bool _oneNoteWasRunning;

    public int MessagesSeen { get; private set; }
    public int SyncEventsParsed { get; private set; }
    /// <summary>Payloads that named a OneNote Storage event but could not be parsed (fail-closed counter).</summary>
    public int MalformedPayloads { get; private set; }
    /// <summary>ETW records the kernel dropped — we may have missed a failure.</summary>
    public int EventsLost { get; private set; }

    public EtwSyncDetector(IEtwMessageSource source, SyncHistoryLog history, NameResolver? names = null,
        Action<WatcherStatus>? onStatusChanged = null, Func<bool>? oneNoteRunning = null, AppLog? log = null,
        TransientPolicy? transient = null, Func<TimeSpan>? awakeClock = null)
    {
        _awakeNow = awakeClock ?? AwakeClock.Stamp;
        _startedAwake = _awakeNow();
        _lastMessageAwake = _startedAwake;
        _source = source; _history = history; _log = log;
        _names = names ?? new NameResolver();
        _transient = transient ?? TransientPolicy.Default;
        _onStatusChanged = onStatusChanged;
        _oneNoteRunning = oneNoteRunning ?? (() => System.Diagnostics.Process.GetProcessesByName("ONENOTE").Length > 0);
    }

    public void Run(CancellationToken ct) => _source.Process(OnMessage, ct);

    /// <summary>Report ETW records dropped by the kernel; surfaced so a gap in coverage is never silent.</summary>
    public void ReportEventsLost(int lost)
    {
        if (lost <= 0) return;
        lock (_gate)
        {
            EventsLost += lost;
            _log?.Warn($"ETW dropped {lost} record(s) (total {EventsLost}) — a sync failure could have been missed");
            var issue = HealthIssues.EventsLost(EventsLost, DateTimeOffset.UtcNow);
            _active[HealthIssues.KeyEventsLost] = new ActiveEntry(issue, HealthIssues.KeyEventsLost, null, issue.FirstSeen);
        }
        _onStatusChanged?.Invoke(Snapshot());
    }

    /// <summary>
    /// System-wide fail-closed sweep, called on a timer.
    ///
    /// The liveness check keys on ANY Office telemetry, not on sync events: measured on real sessions,
    /// an idle OneNote goes up to 114 minutes between sync events, while Office telemetry never stopped
    /// for more than 17 minutes. So "no sync for an hour" is normal; "no telemetry at all" is a dead
    /// pipeline. It is also clamped to our own uptime, so a restart cannot report silence we never watched.
    /// Also reports degraded sources (unreadable index, failing log writes).
    /// </summary>
    public void EvaluateHealth(TimeSpan activityTimeout, DateTimeOffset? nowOverride = null)
    {
        var now = nowOverride ?? DateTimeOffset.UtcNow;
        var changed = false;
        lock (_gate)
        {
            // Liveness, not sync activity: idle OneNote can legitimately go ~2 h without a sync event,
            // but Office telemetry never stops for more than ~17 min on a healthy machine.
            // NEVER claim more silence than we have actually been watching for — after a restart the
            // backfill sets _lastActivity from HISTORICAL data, which is not evidence of live silence.
            // Two things must not be counted as silence we observed: time before this collector started,
            // and time the machine spent asleep. The first is clamped by taking the later of the two
            // stamps; the second by measuring in awake time.
            var startedLater = _lastMessageAt is not { } m || m <= _startedUtc;
            var watchingSince = startedLater ? _startedUtc : _lastMessageAt!.Value;
            var awakeStamp = startedLater ? _startedAwake : _lastMessageAwake;
            // never report more silence than we were actually awake to observe, nor more than the
            // wall clock allows — the smaller of the two can only under-report, never invent an alarm
            var wall = now - watchingSince;
            var awake = _awakeNow() - awakeStamp;
            var since = wall < awake ? wall : awake;
            if (since < TimeSpan.Zero) since = TimeSpan.Zero;
            changed |= _oneNoteRunning() && since > activityTimeout
                ? SetHealth(HealthIssues.KeyStaleActivity, HealthIssues.PipelineSilent(since, activityTimeout, now))
                : RemoveHealth(HealthIssues.KeyStaleActivity);

            var degraded = _names.IndexDegradedReason;
            changed |= degraded is not null
                ? SetHealth(HealthIssues.KeyIndexUnavailable, HealthIssues.IndexUnavailable(degraded, now))
                : RemoveHealth(HealthIssues.KeyIndexUnavailable);

            var writeFails = _history.WriteFailures + (_log?.WriteFailures ?? 0);
            changed |= writeFails > 0
                ? SetHealth(HealthIssues.KeyLogWriteFailing, HealthIssues.LogWriteFailing(writeFails, now))
                : RemoveHealth(HealthIssues.KeyLogWriteFailing);
        }
        if (changed) _onStatusChanged?.Invoke(Snapshot());
    }

    // add/replace only when the message actually changed, so we do not churn status.json
    private bool SetHealth(string key, SyncIssue issue)
    {
        if (_active.TryGetValue(key, out var prev) && prev.Issue.Message == issue.Message) return false;
        var merged = prev is null ? issue : issue with { FirstSeen = prev.Issue.FirstSeen, Occurrences = prev.Issue.Occurrences + 1 };
        _active[key] = new ActiveEntry(merged, key, null, merged.FirstSeen);
        _log?.Warn($"HEALTH {merged.Message} — {merged.Summary}");
        return true;
    }

    private bool RemoveHealth(string key)
    {
        if (!_active.Remove(key)) return false;
        _log?.Info($"health cleared: {key}");
        return true;
    }

    /// <summary>
    /// True exactly once, on the transition from OneNote running to closed. OneNote emits NO final sync
    /// on shutdown (verified on real sessions), so the caller uses this to re-ingest the session log that
    /// has just unlocked and to re-check the cloud side.
    /// </summary>
    public bool OneNoteJustClosed()
    {
        var running = _oneNoteRunning();
        var closed = _oneNoteWasRunning && !running;
        _oneNoteWasRunning = running;
        if (closed) _history.Append($"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}  (OneNote)  CLOSED — no further sync until it is reopened");
        return closed;
    }

    /// <summary>
    /// Restore per-section sync results observed before a restart, from the status file we ourselves
    /// published. Without this every collector restart blinds the cloud check to recent section syncs,
    /// which is enough to resurrect a "change has not reached OneDrive" alert for a section OneNote
    /// already reconciled. Nothing else in the old status is trusted — only these observed facts.
    /// </summary>
    public int SeedSectionSuccesses(IEnumerable<SectionSyncState>? sections)
    {
        if (sections is null) return 0;
        lock (_gate)
        {
            foreach (var x in sections)
                if (!_sectionSuccess.TryGetValue(x.Key, out var had) || x.LastHealthySyncUtc > had.LastHealthySyncUtc)
                    _sectionSuccess[x.Key] = x;
            return _sectionSuccess.Count;
        }
    }

    /// <summary>Add an externally-detected health issue (e.g. an unreadable config) to the published set.</summary>
    public void AddHealthIssue(string key, SyncIssue issue)
    {
        lock (_gate) _active[key] = new ActiveEntry(issue, key, null, issue.FirstSeen);
        _onStatusChanged?.Invoke(Snapshot());
    }

    /// <summary>Feed one record (live source, .etl replay, or diagnostic-log backfill).</summary>
    public void OnMessage(OfficeLogMessage msg)
    {
        MessagesSeen++;
        _lastMessageAt = DateTimeOffset.UtcNow;
        _lastMessageAwake = _awakeNow();
        if (!SyncEventJson.LooksLikeSyncSendEvent(msg.Message)) return;
        var ev = SyncEventJson.TryParse(msg.Message, out var malformed);
        if (malformed) MalformedPayloads++;
        if (ev is null) return;
        SyncEventsParsed++;
        Ingest(ev);
    }

    public void Ingest(SyncEvent ev)
    {
        _names.Refresh(TimeSpan.FromMinutes(2));
        var names = _names.Resolve(ev);
        var changed = false;
        var t = ev.Time == default ? DateTimeOffset.UtcNow : ev.Time;

        lock (_gate)
        {
            _history.Append(SyncEventMapper.ToHistoryLine(ev, names));

            // an event we have never classified: log it once, loudly, so it can be catalogued
            if (ev.Outcome is SyncOutcome.Unknown or SyncOutcome.SuspectedFailure && _seenUnknownEvents.Add(ev.EventName))
                _log?.Warn($"UNCLASSIFIED EVENT '{ev.EventName}' → treated as a possible failure ({ev.ClassificationReason}). Please report it.");

            if (ev.Kind == SyncEventKind.ConnectivityChanged)
            {
                var was = _internet; _internet = ev.InternetAvailable;
                if (was != _internet)
                {
                    _log?.Info($"connectivity: {(_internet == true ? "online" : "OFFLINE")}");
                    changed = UpdateOfflineIssue();
                }
            }
            else
            {
                var (scopeKey, sectionRid) = SyncScope.Of(ev, names);

                if (ev.IsProblem)
                {
                    // a failure already covered by a LATER proven success must not re-open the error
                    // (the post-exit backfill replays finished session logs, so this really happens)
                    if (_coveredUntil.TryGetValue(scopeKey, out var coveredAt) && coveredAt >= t)
                    {
                        _log?.Info($"not raising {scopeKey}: a later success at {coveredAt.ToLocalTime():HH:mm:ss} already covers it");
                    }
                    else if (ev.Outcome == SyncOutcome.Transient && _transient.ShouldHoldBack(scopeKey, ev.ErrorCodeHex, t))
                    {
                        // the only deliberate hold-back is a code the user listed in [transient]
                        _log?.Info($"held back by [transient] policy: {ev.ErrorCodeHex} on {scopeKey}");
                    }
                    else
                    {
                        var issue = SyncEventMapper.ToIssue(ev, DetectorName, names);
                        if (issue is not null)
                        {
                            if (_active.TryGetValue(scopeKey, out var prev))
                                issue = issue with { FirstSeen = prev.Issue.FirstSeen, Occurrences = prev.Issue.Occurrences + 1 };
                            _active[scopeKey] = new ActiveEntry(issue, scopeKey, sectionRid, t);
                            changed = true;
                            // Put the notebook on the board even though this is a failure. Previously
                            // _notebooks was written only on success, so a notebook whose only event was
                            // an error vanished from the NOTEBOOKS table — a fail-closed app hiding the
                            // very thing it exists to show.
                            TouchNotebook(ev, names, t);
                            var where = issue.Location.Length > 0 ? issue.Location : "<" + ev.EventName + ">";
                            _log?.Warn($"ISSUE {where}  {issue.Code} {issue.Message}  [{issue.Category}] x{issue.Occurrences}  ({ev.Outcome}: {ev.ClassificationReason})");
                        }
                    }
                }
                else if (ev.Outcome == SyncOutcome.Success)
                {
                    changed |= ClearCoveredBy(ev, names, t);

                    // Any healthy section-scoped activity counts, not just a full section sync. This
                    // record answers "why did the local timestamp move?", and OneNote re-stamps the
                    // search index on ANY reconciliation — a real-time round trip or a page transfer
                    // included. Measured 2026-09-06: a section can go hours with healthy real-time and
                    // notebook activity and no SectionSyncResult at all, so requiring one made the
                    // answer unobtainable and left a false alert standing for over four hours.
                    if (ev.Kind is SyncEventKind.SectionSyncResult or SyncEventKind.RealTimeService
                                or SyncEventKind.PageUpload or SyncEventKind.PageDownload
                        && SectionKey.Normalize(ev.SectionResourceId ?? ev.SectionGosid ?? names.Section) is { } sk)
                    {
                        if (!_sectionSuccess.TryGetValue(sk, out var had) || t > had.LastHealthySyncUtc)
                        {
                            _sectionSuccess[sk] = new SectionSyncState(sk, names.Section ?? sk, t);
                            changed = true;
                        }
                    }

                    if (ev.IsSyncActivity)
                    {
                        _lastActivity = t; changed = true;
                        TouchNotebook(ev, names, t);
                    }
                }
            }
        }

        if (changed) _onStatusChanged?.Invoke(Snapshot());
    }

    /// <summary>
    /// Clear only the errors this success genuinely COVERS (see <see cref="SyncScope"/>) and only when
    /// the success is LATER than the failure. Anything else stays an error — one we cannot prove
    /// resolved is still a problem.
    /// </summary>
    private bool ClearCoveredBy(SyncEvent ev, ResolvedNames names, DateTimeOffset t)
    {
        var covered = SyncScope.CoveredBy(ev, names);
        if (covered.Count == 0) return false;

        foreach (var key in covered)
            if (!_coveredUntil.TryGetValue(key, out var prev) || t > prev) _coveredUntil[key] = t;

        var sectionRid = ev.SectionResourceId ?? ev.SectionGosid ?? names.Section;
        var clearable = _active.Where(kv =>
        {
            if (kv.Value.FailedAt >= t) return false;                     // the success predates the failure
            if (covered.Contains(kv.Key)) return true;                    // exact scope
            return ev.Kind == SyncEventKind.SectionSyncResult && sectionRid is not null
                   && SyncScope.SectionCoversPage(kv.Key, kv.Value.SectionRid, sectionRid);
        }).Select(kv => kv.Key).ToList();

        foreach (var key in clearable)
        {
            _active.Remove(key);
            _transient.Forget(key);
            _log?.Info($"RECOVERED {key} — covered by a later {ev.Kind} success at {t.ToLocalTime():HH:mm:ss}");
        }
        return clearable.Count > 0;
    }

    private bool UpdateOfflineIssue()
    {
        const string key = "network:offline";
        if (_internet == false)
        {
            var offline = new SyncIssue
            {
                Detector = DetectorName, Kind = IssueKind.Offline,
                Category = Core.Diagnosis.FailureCategory.Network,
                Message = "No internet connectivity",
                Summary = "OneNote reports the PC is offline; nothing can sync until connectivity returns.",
                Recommendation = "Reconnect to Wi-Fi/VPN. Sync resumes automatically.",
                FirstSeen = DateTimeOffset.UtcNow, LastSeen = DateTimeOffset.UtcNow,
            };
            _active[key] = new ActiveEntry(offline, key, null, offline.FirstSeen);
            return true;
        }
        return _active.Remove(key);
    }

    /// <summary>Notebook key used by the status table — the resolved name when we have one.</summary>
    private static string NotebookKeyOf(SyncEvent ev, ResolvedNames names) =>
        names.Notebook ?? ev.NotebookGosid ?? ev.NotebookResourceId ?? "?";

    /// <summary>
    /// Record that we saw activity for this event's notebook. Called for successes AND failures: the
    /// table stores last-known activity, and whether the notebook is currently healthy is decided at
    /// publish time from the live issue set (see <see cref="Snapshot"/>), so it can never go stale.
    /// </summary>
    private void TouchNotebook(SyncEvent ev, ResolvedNames names, DateTimeOffset t)
    {
        var key = NotebookKeyOf(ev, names);
        if (key == "?") return;
        var prev = _notebooks.GetValueOrDefault(key);
        var last = prev?.LastSyncUtc is { } p && p > t ? p : t;
        _notebooks[key] = new NotebookStatus(names.Notebook ?? key, ev.NotebookGosid ?? prev?.Gosid, last,
            true, null, null, names.NotebookDisplayName ?? prev?.DisplayName);
    }

    public WatcherStatus Snapshot()
    {
        lock (_gate)
        {
            var now = DateTimeOffset.UtcNow;
            foreach (var k in _active.Where(kv => kv.Value.Issue.IsExpired(now)).Select(kv => kv.Key).ToList())
            {
                _active.Remove(k);
                _log?.Info($"issue expired (not seen again): {k}");
            }
            return new WatcherStatus
            {
                CollectorRunning = true,
                OneNoteRunning = _oneNoteRunning(),
                InternetAvailable = _internet,
                LastSyncEventUtc = _lastActivity,
                LastTelemetryUtc = _lastMessageAt,
                Notebooks = PublishNotebooks(),
                Sections = _sectionSuccess.Values.OrderBy(x => x.Name, StringComparer.OrdinalIgnoreCase).ToList(),
                ActiveIssues = _active.Values.Select(v => v.Issue).ToList(),
                EventsLost = EventsLost,
                MalformedPayloads = MalformedPayloads,
            };
        }
    }

    /// <summary>
    /// The notebook table as published. A notebook carrying an open issue is reported FAILED with that
    /// issue's code and message; only a notebook with nothing outstanding reads OK. Derived rather than
    /// stored so the table can never disagree with the PROBLEMS list. A notebook that has only ever
    /// failed still gets a row.
    /// </summary>
    private List<NotebookStatus> PublishNotebooks()
    {
        var failures = new Dictionary<string, SyncIssue>(StringComparer.OrdinalIgnoreCase);
        foreach (var e in _active.Values)
        {
            var nb = e.Issue.NotebookName;
            if (nb is null) continue;
            if (!failures.TryGetValue(nb, out var prev) || e.Issue.LastSeen > prev.LastSeen) failures[nb] = e.Issue;
        }

        var rows = _notebooks.ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.OrdinalIgnoreCase);
        foreach (var (nb, issue) in failures)
        {
            var prev = rows.GetValueOrDefault(nb);
            rows[nb] = new NotebookStatus(prev?.Name ?? nb, prev?.Gosid, prev?.LastSyncUtc ?? issue.LastSeen,
                false, issue.Code, issue.Message, prev?.DisplayName ?? issue.NotebookDisplayName);
        }
        return rows.Values.OrderBy(n => n.Name, StringComparer.OrdinalIgnoreCase).ToList();
    }

    public IReadOnlyCollection<SyncIssue> ActiveIssues()
    {
        lock (_gate) return _active.Values.Select(v => v.Issue).ToList();
    }
}
