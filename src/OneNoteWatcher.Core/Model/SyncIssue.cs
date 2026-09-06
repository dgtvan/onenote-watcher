using OneNoteWatcher.Core.Diagnosis;

namespace OneNoteWatcher.Core.Model;

public enum IssueKind
{
    SyncFailed,
    UploadStuck,
    DownloadStuck,
    ErrorDialog,
    ErrorCodeReported,
    AuthRequired,
    SourceUnavailable,
    Offline,
}

/// <summary>
/// Something is wrong and you need to look at it.
///
/// There is deliberately NO severity: the watcher has exactly two states — error and success. A
/// middle "warning" tier is the thing people learn to ignore, which is precisely how a rare sync
/// failure gets missed. If an issue exists and is not ignored by config, the tray is red.
/// Noise is controlled explicitly instead: <c>[ignore]</c> rules, the <c>[transient]</c> policy, and
/// <see cref="ExpiresUtc"/> for signals that have no natural "cleared" event.
///
/// Every issue answers: WHERE (notebook / section / page), WHAT (code + symbolic name),
/// WHY (<see cref="Category"/> + <see cref="Summary"/>) and HOW TO FIX (<see cref="Recommendation"/>).
/// </summary>
public sealed record SyncIssue
{
    public required string Detector { get; init; }
    public required IssueKind Kind { get; init; }

    public string? NotebookName { get; init; }
    /// <summary>The nickname set in OneNote's navigation pane when it differs from <see cref="NotebookName"/>.</summary>
    public string? NotebookDisplayName { get; init; }
    public string? SectionName { get; init; }
    /// <summary>The section's OneDrive/Graph resource id when known, so a finding can be matched to the
    /// collector's per-section results across the several id spellings (see <see cref="SectionKey"/>).</summary>
    public string? SectionId { get; init; }
    public string? PageTitle { get; init; }

    /// <summary>The OneNote telemetry event this came from, so <c>[ignore] events</c> can target it.</summary>
    public string? EventName { get; init; }

    /// <summary>"0xE000005D" style code when known.</summary>
    public string? Code { get; init; }

    /// <summary>
    /// Short machine-ish message, e.g. the symbolic description (ErrFilePendingRename). Never null —
    /// an issue nobody can describe is not actionable — so callers and rules can use it directly.
    /// </summary>
    public string Message { get; init; } = "";

    public FailureCategory Category { get; init; } = FailureCategory.Unknown;

    /// <summary>Plain-language explanation of what went wrong and why.</summary>
    public string? Summary { get; init; }

    /// <summary>Concrete steps to fix it.</summary>
    public string? Recommendation { get; init; }

    /// <summary>Raw ids for correlation (section resource id, notebook GOSID…) when names are unknown.</summary>
    public string? TechnicalDetail { get; init; }

    public string? Reference { get; init; }

    public DateTimeOffset FirstSeen { get; init; }
    public DateTimeOffset LastSeen { get; init; }
    public int Occurrences { get; init; } = 1;

    /// <summary>For inferred issues: the timestamp of the evidence (e.g. the local change that is not
    /// reaching the cloud). A later successful sync of the same scope from the ETW collector refutes it.</summary>
    public DateTimeOffset? EvidenceUtc { get; init; }

    /// <summary>
    /// When set, the issue disappears at this time unless seen again. Used for signals that have no
    /// natural "cleared" event (a suspected/unknown failure), so they cannot stay red forever — the
    /// lesson from the SyncScore false alarm. Confirmed failures have no expiry: they clear only on a
    /// real success for the same scope.
    /// </summary>
    public DateTimeOffset? ExpiresUtc { get; init; }

    public bool IsExpired(DateTimeOffset now) => ExpiresUtc is { } e && now > e;

    /// <summary>Dedupe key: same detector+kind+scope+code collapses into one live issue.</summary>
    public string DedupeKey =>
        string.Join('|', Detector, Kind, NotebookName ?? "", SectionName ?? "", Code ?? "");

    /// <summary>"Note (Display name: Van) / ASW / Eyes check" style location string.</summary>
    public string Location
    {
        get
        {
            var nb = NotebookName is null ? null
                : NotebookDisplayName is null ? NotebookName : $"{NotebookName} (Display name: {NotebookDisplayName})";
            return string.Join(" / ", new[] { nb, SectionName, PageTitle }.Where(s => !string.IsNullOrEmpty(s)));
        }
    }

    /// <summary>Multi-line human report for balloons, the issues window and the history.</summary>
    public string ToReport()
    {
        var lines = new List<string>
        {
            $"WHERE : {(Location.Length > 0 ? Location : "(unknown location)")}",
            $"WHAT  : {Kind}{(Code is null ? "" : $"  {Code}")}{(Message.Length == 0 ? "" : $"  {Message}")}",
            $"WHY   : [{Category}] {Summary}",
            $"FIX   : {Recommendation}",
        };
        if (TechnicalDetail is not null) lines.Add($"IDS   : {TechnicalDetail}");
        if (Reference is not null) lines.Add($"MORE  : {Reference}");
        lines.Add($"SEEN  : first {FirstSeen.ToLocalTime():yyyy-MM-dd HH:mm}, last {LastSeen.ToLocalTime():HH:mm}, {Occurrences}×");
        return string.Join(Environment.NewLine, lines);
    }
}
