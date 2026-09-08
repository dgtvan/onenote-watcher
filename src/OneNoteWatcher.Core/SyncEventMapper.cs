using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core;

/// <summary>Turns a classified <see cref="SyncEvent"/> into a history line and (on a problem) an explained issue.</summary>
public static class SyncEventMapper
{
    /// <summary>How long a non-confirmed problem stays visible without being seen again.</summary>
    public static readonly TimeSpan SuspectedIssueTtl = TimeSpan.FromHours(6);

    public static string ToHistoryLine(SyncEvent e, ResolvedNames? names = null)
    {
        names ??= new ResolvedNames(null, null, null);
        var when = (e.Time == default ? DateTimeOffset.Now : e.Time.ToLocalTime()).ToString("yyyy-MM-dd HH:mm:ss zzz");

        if (e.Kind == SyncEventKind.ConnectivityChanged)
            return $"{when}  (network)  {(e.InternetAvailable == true ? "ONLINE" : "OFFLINE")}";

        var nb = names.NotebookLabel ?? e.NotebookGosid ?? e.NotebookResourceId ?? "?";
        var scope = e.Kind switch
        {
            SyncEventKind.SectionSyncResult or SyncEventKind.PageUpload or SyncEventKind.PageDownload or SyncEventKind.RealTimeService
                => $"{nb} / {names.Section ?? ShortId(e.SectionResourceId)}",
            SyncEventKind.PageSyncSession => $"{nb} / {names.Section ?? "?"} / {names.Page ?? ShortId(e.PageGoid)}",
            SyncEventKind.SyncScore => "(replication scan)",
            SyncEventKind.UnrecognisedStorage or SyncEventKind.OtherOneNoteSignal
                => names.Notebook is null && e.SectionResourceId is null ? "(no location)" : $"{nb} / {names.Section ?? ShortId(e.SectionResourceId)}",
            _ => nb,
        };

        var verb = e.Outcome switch
        {
            SyncOutcome.Success => e.Kind switch
            {
                SyncEventKind.PageUpload => $"PAGE-UPLOAD  OK  ({e.TransferTimeMs} ms)",
                SyncEventKind.PageDownload => $"PAGE-DOWNLOAD  OK  ({e.TransferTimeMs} ms)",
                SyncEventKind.RealTimeService => "REALTIME  OK",
                _ => "OK",
            },
            // keep the channel visible on failures too, so the history says WHAT failed
            SyncOutcome.Failure => e.Kind == SyncEventKind.RealTimeService ? "REALTIME  FAILED" : "FAILED",
            SyncOutcome.Transient => e.Kind == SyncEventKind.RealTimeService ? "REALTIME  FAILED(transient)" : "FAILED(transient)",
            SyncOutcome.SuspectedFailure => e.Kind == SyncEventKind.PageSyncSession
                ? $"PAGE-SESSION  SUSPECT  errorStateMs={e.TimeInSyncErrorStateMs}" : "SUSPECT",
            SyncOutcome.Unknown => "UNKNOWN",
            _ => e.Kind switch
            {
                SyncEventKind.SyncScore => "SCORE",
                SyncEventKind.PageSyncSession => $"PAGE-SESSION  OK  ({e.TimeInSyncErrorStateMs} ms in error state)",
                _ => "INFO",
            },
        };

        var code = e.ErrorCodeHex is { } hex ? $"  {hex} {e.ErrorDescription ?? e.ErrorType ?? ""}".TrimEnd()
                 : e.ErrorDescription is { } d ? $"  {d}" : "";
        var why = "";
        if (e.Outcome is SyncOutcome.Failure or SyncOutcome.Transient)
        {
            var ex = ErrorCatalog.Explain(e.ErrorCode, e.ErrorDescription, e.ErrorType);
            why = $"  [{ex.Category}] {ex.Summary}";
        }
        else if (e.Outcome is SyncOutcome.SuspectedFailure or SyncOutcome.Unknown or SyncOutcome.Diagnostic)
            why = e.ClassificationReason is { } r ? $"  ({r})" : "";

        var evName = e.Kind is SyncEventKind.UnrecognisedStorage or SyncEventKind.OtherOneNoteSignal or SyncEventKind.SyncScore
            ? $"  <{e.EventName}>" : "";
        return $"{when}  {scope}  {verb}{code}{evName}{why}";
    }

    /// <summary>
    /// An explained issue for anything that is not a proven success — the watcher has two states, so any
    /// issue means the tray is red. Unproven signals carry a TTL so they self-clear without evidence.
    /// Diagnostic events never produce an issue.
    /// </summary>
    public static SyncIssue? ToIssue(SyncEvent e, string detector, ResolvedNames? names = null)
    {
        if (!e.IsProblem) return null;
        names ??= new ResolvedNames(null, null, null);

        var ex = ErrorCatalog.Explain(e.ErrorCode, e.ErrorDescription, e.ErrorType);
        var now = e.Time == default ? DateTimeOffset.Now : e.Time;
        var realtime = e.Kind == SyncEventKind.RealTimeService;

        // "No TTL" means "held until a success proves it resolved" — which requires that such a success
        // can exist. For these two kinds it cannot: SyncScope keys them by event NAME, and CoveredBy
        // never returns an event key, so nothing in the system is able to clear one. Left without a TTL
        // they would not be held pending proof, they would be stuck red forever. Time is their only
        // exit, so they get it whatever the outcome — and a problem that is still happening re-arms the
        // TTL on every occurrence, so only one that genuinely stopped fades.
        var clearedOnlyByTime = e.Kind is SyncEventKind.UnrecognisedStorage or SyncEventKind.OtherOneNoteSignal;
        var ttl = (DateTimeOffset?)(now + SuspectedIssueTtl);

        var (kind, expires) = e.Outcome switch
        {
            SyncOutcome.Failure => (
                e.ErrorCodeHex is not null ? IssueKind.ErrorCodeReported : IssueKind.SyncFailed,
                clearedOnlyByTime ? ttl : null),
            SyncOutcome.Transient => (IssueKind.ErrorCodeReported, clearedOnlyByTime ? ttl : null),
            // an unproven signal self-clears, so it cannot stay red forever without evidence
            _ => (IssueKind.SyncFailed, ttl),   // SuspectedFailure / Unknown
        };

        var (summary, recommendation) = e.Outcome switch
        {
            SyncOutcome.Failure or SyncOutcome.Transient when realtime =>
                ($"The real-time page sync channel reported: {e.ErrorDescription}",
                 "Usually self-heals within a minute. If it persists, press Shift+F9 in OneNote and check your connection."),
            SyncOutcome.Failure or SyncOutcome.Transient =>
                (ex.Summary + (e.Outcome == SyncOutcome.Transient ? "  OneNote flagged this as retryable, but it has not recovered." : ""),
                 ex.Recommendation),
            SyncOutcome.SuspectedFailure =>
                ($"OneNote emitted '{e.EventName}', which indicates a sync problem, but without a definite outcome. Reason: {e.ClassificationReason}.",
                 "Check File → Info → View Sync Status in OneNote for this notebook. If sync is fine, this clears itself; the event is recorded in the history for reference."),
            _ =>
                ($"OneNote emitted '{e.EventName}', which this watcher does not recognise and cannot confirm as successful. Reason: {e.ClassificationReason}.",
                 "Treated as a possible failure so it is not missed. Check File → Info → View Sync Status; if sync is healthy this clears itself. Please report the event name so it can be classified."),
        };

        var ids = new List<string>();
        if (e.NotebookGosid is not null) ids.Add($"notebook-gosid={e.NotebookGosid}");
        if (e.NotebookResourceId is not null) ids.Add($"notebook-rid={e.NotebookResourceId}");
        if (e.SectionResourceId is not null) ids.Add($"section-rid={e.SectionResourceId}");
        if (e.SectionGosid is not null) ids.Add($"section-gosid={e.SectionGosid}");
        ids.Add($"event={e.EventName}");
        foreach (var kv in e.ErrorFields) ids.Add($"{kv.Key[5..]}={kv.Value}");

        return new SyncIssue
        {
            Detector = detector,
            Kind = kind,
            NotebookName = names.Notebook ?? e.NotebookGosid ?? e.NotebookResourceId,
            NotebookDisplayName = names.NotebookDisplayName,
            SectionName = names.Section ?? (e.SectionResourceId is not null ? ShortId(e.SectionResourceId) : null),
            PageTitle = names.Page,
            EventName = e.EventName,
            Code = e.ErrorCodeHex,
            Message = e.ErrorDescription ?? e.ErrorType ?? (e.Outcome == SyncOutcome.Unknown ? "Unclassified OneNote event" : "Sync failed"),
            Category = ex.Category,
            Summary = summary,
            Recommendation = recommendation,
            Reference = ex.Reference,
            TechnicalDetail = string.Join("; ", ids),
            FirstSeen = now, LastSeen = now,
            ExpiresUtc = expires,
        };
    }

    private static string ShortId(string? id) => id is null ? "?" : id.Length > 14 ? "…" + id[^12..] : id;
}
