using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Health;

/// <summary>
/// System-wide fail-closed rules: the watcher must never look healthy when it cannot actually see.
///
/// A missed sync failure is the catastrophic outcome; an extra warning is cheap. So every condition
/// that degrades the watcher's ability to observe — a dead ETW pipeline, an unreadable index, a blind
/// cloud check, logs that will not write — becomes an ERROR instead of silent green. There is no
/// middle 'warning' tier: two states only, error and success.
/// Each factory produces a stable <c>Key</c> so the issue can be added/removed idempotently.
/// </summary>
public static class HealthIssues
{
    public const string DetectorName = "health";

    private static SyncIssue Make(string key, IssueKind kind, FailureCategory cat,
        string message, string summary, string recommendation, DateTimeOffset now, DateTimeOffset? expires = null) => new()
    {
        Detector = DetectorName, Kind = kind, Category = cat,
        Message = message, Summary = summary, Recommendation = recommendation,
        TechnicalDetail = $"health-key={key}",
        FirstSeen = now, LastSeen = now, ExpiresUtc = expires,
    };

    /// <summary>Keys are stable so callers can add/remove without duplicating.</summary>
    public const string KeyStaleActivity = "health:stale-activity";
    public const string KeyCollectorDown = "health:collector-down";
    public const string KeyIndexUnavailable = "health:index-unavailable";
    public const string KeyGraphBlind = "health:graph-blind";
    public const string KeyOAlertsUnavailable = "health:oalerts-unavailable";
    public const string KeyLogWriteFailing = "health:log-write-failing";
    public const string KeyDetectorDisabled = "health:detector-disabled";
    public const string KeyConfigUnreadable = "health:config-unreadable";
    public const string KeyEventsLost = "health:events-lost";

    /// <summary>
    /// THE most important one: OneNote is running but the watcher is receiving NO Office telemetry at
    /// all, so its pipeline is dead and a sync failure would go unnoticed.
    ///
    /// This deliberately keys on ANY telemetry, not on sync events. Measured on real sessions: gaps
    /// between OneNote *sync* events reach 114 minutes when OneNote is simply idle (median 0 m, p90
    /// 15 m), so "no sync for an hour" is normal and alerting on it is a false alarm. Gaps between
    /// *any* Office telemetry never exceeded 17 minutes — so silence there really does mean the
    /// pipeline stopped.
    /// </summary>
    public static SyncIssue PipelineSilent(TimeSpan since, TimeSpan threshold, DateTimeOffset now) => Make(
        KeyStaleActivity, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        $"no telemetry received for {Human(since)}",
        $"OneNote is running but the watcher has received no Office telemetry at all for {Human(since)} " +
        $"(never more than ~17 minutes on a healthy machine, threshold {Human(threshold)}). Its ETW pipeline " +
        "has most likely stopped, which means a sync failure would go unnoticed.",
        @"Restart the 'OneNoteWatcher Collector' scheduled task (or re-run scripts\install.ps1 as Administrator), " +
        "then check OneNote's own File → Info → View Sync Status to confirm nothing failed meanwhile.",
        now);

    public static SyncIssue CollectorDown(DateTimeOffset? lastSeen, DateTimeOffset now) => Make(
        KeyCollectorDown, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        "real-time collector is not running",
        "The elevated collector is not publishing status, so real-time sync failures are not being detected. " +
        (lastSeen is { } t ? $"Its last update was {t.ToLocalTime():yyyy-MM-dd HH:mm}." : "It has never reported."),
        "Run scripts\\install.ps1 as Administrator, or start the 'OneNoteWatcher Collector' scheduled task.",
        now);

    public static SyncIssue IndexUnavailable(string reason, DateTimeOffset now) => Make(
        KeyIndexUnavailable, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        "OneNote's local index could not be read",
        $"The watcher could not read OneNote's local search index ({reason}), so notebook/section names and the " +
        "cloud-side comparison may be incomplete.",
        "This usually resolves itself. If it persists, confirm OneNote is installed and has been opened at least once.",
        now);

    /// <summary>
    /// The Microsoft sign-in has EXPIRED. The cloud check is the only
    /// thing that can prove your edits reached OneDrive (nothing else can, especially after OneNote is
    /// closed), so an expired token means the watcher has gone blind to the worst failure mode.
    /// </summary>
    public static SyncIssue GraphTokenExpired(DateTimeOffset now) => Make(
        KeyGraphBlind, IssueKind.AuthRequired, FailureCategory.Permission,
        "Microsoft sign-in has expired",
        "The watcher's Microsoft sign-in has expired, so it can no longer check whether your notes actually " +
        "reached OneDrive. Sync failures that leave changes stranded on this PC would NOT be detected until you sign in again.",
        "Open 'Show issues & status' and click 'Sign in to Graph…' to sign in again (one device code, takes a minute).",
        now);

    public static SyncIssue GraphBlind(string reason, DateTimeOffset now) => Make(
        KeyGraphBlind, IssueKind.AuthRequired, FailureCategory.Permission,
        "cloud-side check unavailable",
        $"The Microsoft Graph check is not running ({reason}). Without it the watcher cannot confirm that your local " +
        "changes actually reached OneDrive, and section names may be missing from failures.",
        "Open 'Show issues & status' and click 'Sign in to Graph…'. If already signed in, check the tray log for the error.",
        now);

    public static SyncIssue OAlertsUnavailable(DateTimeOffset now) => Make(
        KeyOAlertsUnavailable, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        "OneNote error dialogs are not being watched",
        "The watcher could not subscribe to the Windows OAlerts event log, so error dialogs OneNote shows will not be detected.",
        "Confirm the 'OAlerts' event log exists and is enabled (it ships with Office).",
        now);

    public static SyncIssue LogWriteFailing(int failures, DateTimeOffset now) => Make(
        KeyLogWriteFailing, IssueKind.SourceUnavailable, FailureCategory.Storage,
        $"{failures} log write(s) failed",
        "The watcher cannot write its history/log files, so a record of sync results is being lost.",
        "Check free disk space and that your user can write to the watcher's folder (C:\\ProgramData\\OneNoteWatcher\\logs).",
        now);

    public static SyncIssue DetectorDisabled(string what, DateTimeOffset now) => Make(
        KeyDetectorDisabled + ":" + what, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        $"{what} is disabled in config",
        $"{what} is switched off in config.ini, so that part of sync monitoring is not running.",
        "Set it back to enabled = true in config.ini if you did not intend this.",
        now);

    public static SyncIssue ConfigUnreadable(string path, string reason, DateTimeOffset now) => Make(
        KeyConfigUnreadable, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        "config.ini could not be read",
        $"The watcher could not read {path} ({reason}), so it is running on defaults — your ignore rules are NOT applied.",
        "Fix or restore config.ini, then restart the watcher.",
        now);

    public static SyncIssue EventsLost(int lost, DateTimeOffset now) => Make(
        KeyEventsLost, IssueKind.SourceUnavailable, FailureCategory.Unknown,
        $"{lost} telemetry record(s) dropped",
        "Windows dropped ETW records under load, so a sync event may not have been seen by the watcher.",
        "Check OneNote's own sync status (File → Info → View Sync Status) to confirm nothing failed during this period.",
        now, expires: now.AddHours(6));

    private static string Human(TimeSpan t) =>
        t.TotalHours >= 1 ? $"{t.TotalHours:F1} h" : $"{t.TotalMinutes:F0} min";
}
