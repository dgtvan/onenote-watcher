using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Index;
using System.Text.Json;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Detection;

/// <summary>A section as seen in the local search index: newest LastModifiedTime in its subtree.</summary>
public sealed record LocalSection(string? Notebook, string Section, DateTimeOffset? Newest);

/// <summary>
/// BASELINE, no-admin, cloud-side check — STATEFUL. Verified on 2026-09-06 that the search index's
/// LastModifiedTime is re-stamped when a notebook is (re)opened or re-synced, so it is NOT "last content
/// edit" and cannot be compared to Graph's lastModifiedDateTime absolutely (that produced 15 false
/// "upload stuck" alerts on a healthy machine). Instead: on the first poll every section is baselined;
/// an UploadStuck is raised only when a section's local timestamp advances *while watching* and the
/// server's timestamp then fails to advance within the grace period. Server-only movement is reported
/// only when it persists beyond the grace period. When the ETW collector is
/// running its per-sync results are authoritative and the tray suppresses outcome issues it contradicts.
/// </summary>
public sealed class OutcomeDetector
{
    public const string DetectorName = "outcome";
    private static readonly TimeSpan SyncTolerance = TimeSpan.FromMinutes(2);

    private readonly TimeSpan _grace;
    private readonly Dictionary<string, (DateTimeOffset local, DateTimeOffset? server)> _baseline = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, DateTimeOffset> _aheadSince = new(StringComparer.OrdinalIgnoreCase);

    public OutcomeDetector(TimeSpan grace) => _grace = grace;

    public int BaselinedSections => _baseline.Count;

    public static IReadOnlyList<LocalSection> LocalSectionsFrom(SearchIndexReader index) =>
        index.Sections().Select(s => new LocalSection(index.NotebookNameOf(s), s.Title ?? "", index.NewestIn(s))).ToList();

    public IReadOnlyList<SyncIssue> Evaluate(SearchIndexReader index, GraphSnapshot server, bool oneNoteRunning, bool? internet, DateTimeOffset? now = null)
        => Evaluate(LocalSectionsFrom(index), server, oneNoteRunning, internet, now);

    public IReadOnlyList<SyncIssue> Evaluate(IReadOnlyList<LocalSection> local, GraphSnapshot server, bool oneNoteRunning, bool? internet, DateTimeOffset? now = null)
    {
        var at = now ?? DateTimeOffset.UtcNow;
        var issues = new List<SyncIssue>();

        foreach (var gs in server.Sections)
        {
            var match = local.FirstOrDefault(s => Eq(s.Section, gs.DisplayName) && Eq(s.Notebook, gs.NotebookName))
                        ?? local.FirstOrDefault(s => Eq(s.Section, gs.DisplayName));
            if (match?.Newest is null || gs.LastModified is null) continue;

            var key = $"{gs.NotebookName}/{gs.DisplayName}";
            var localNewest = match.Newest.Value;
            var serverNewest = gs.LastModified.Value;

            if (!_baseline.TryGetValue(key, out var b))
            {
                _baseline[key] = (localNewest, serverNewest);   // first sighting: learn, never judge
                continue;
            }

            var localMoved = localNewest > b.local;
            var serverMoved = serverNewest > (b.server ?? DateTimeOffset.MinValue);
            var serverCaughtUp = serverNewest >= localNewest - SyncTolerance;

            if (localMoved && (serverCaughtUp || serverMoved))
            {
                _baseline[key] = (localNewest, serverNewest); _aheadSince.Remove(key);   // synced
                continue;
            }
            if (localMoved && !serverCaughtUp)
            {
                if (!_aheadSince.TryGetValue(key, out var since)) { _aheadSince[key] = since = at; }
                if (at - since >= _grace)
                {
                    var ex = ErrorCatalog.UploadStuck(oneNoteRunning, internet);
                    issues.Add(new SyncIssue
                    {
                        Detector = DetectorName, Kind = IssueKind.UploadStuck,
                        NotebookName = match.Notebook ?? gs.NotebookName, SectionName = gs.DisplayName,
                        Message = $"a local change at {localNewest.ToLocalTime():HH:mm} has not reached OneDrive after {Human(at - since)}",
                        Category = ex.Category, Summary = ex.Summary, Recommendation = ex.Recommendation,
                        TechnicalDetail = $"local={localNewest:u} server={serverNewest:u} graph-section={gs.Id}",
                        EvidenceUtc = localNewest,
                        FirstSeen = since, LastSeen = at,
                    });
                }
                continue;
            }
            if (!localMoved && serverMoved && oneNoteRunning && serverNewest - localNewest > _grace)
            {
                _baseline[key] = (localNewest, serverNewest);
                issues.Add(new SyncIssue
                {
                    Detector = DetectorName, Kind = IssueKind.DownloadStuck,
                    NotebookName = match.Notebook ?? gs.NotebookName, SectionName = gs.DisplayName,
                    Message = $"OneDrive changed at {serverNewest.ToLocalTime():HH:mm}; not yet reflected locally",
                    Category = FailureCategory.Unknown,
                    Summary = "A change made elsewhere has not arrived on this PC yet (informational).",
                    Recommendation = "Nothing to do unless it persists; then press Shift+F9 in OneNote.",
                    TechnicalDetail = $"local={localNewest:u} server={serverNewest:u}",
                    EvidenceUtc = serverNewest,
                    FirstSeen = at, LastSeen = at,
                });
                continue;
            }
            if (!localMoved) _baseline[key] = (b.local, serverNewest); // keep server side current
        }
        return issues;
    }

    /// <summary>
    /// Persisted so an edit that never reached the cloud is still flagged after a reboot. Without this,
    /// "edit → close OneNote → shut down" would re-baseline on next login and the stranded change would
    /// be silently accepted as normal.
    /// </summary>
    private sealed record PersistedState(
        Dictionary<string, DateTimeOffset> BaselineLocal,
        Dictionary<string, DateTimeOffset?> BaselineServer,
        Dictionary<string, DateTimeOffset> AheadSince);

    public void SaveState(string path)
    {
        try
        {
            var state = new PersistedState(
                _baseline.ToDictionary(k => k.Key, v => v.Value.local),
                _baseline.ToDictionary(k => k.Key, v => v.Value.server),
                new Dictionary<string, DateTimeOffset>(_aheadSince));
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(state));
            File.Move(tmp, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { }
    }

    /// <summary>Returns the number of sections restored, or 0 when there was nothing to restore.</summary>
    public int LoadState(string path)
    {
        try
        {
            if (!File.Exists(path)) return 0;
            var state = JsonSerializer.Deserialize<PersistedState>(File.ReadAllText(path));
            if (state is null) return 0;
            _baseline.Clear(); _aheadSince.Clear();
            foreach (var kv in state.BaselineLocal)
                _baseline[kv.Key] = (kv.Value, state.BaselineServer.GetValueOrDefault(kv.Key));
            foreach (var kv in state.AheadSince) _aheadSince[kv.Key] = kv.Value;
            return _baseline.Count;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return 0; }
    }

    private static bool Eq(string? a, string? b) => string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    private static string Human(TimeSpan t) => t.TotalHours >= 1 ? $"{t.TotalHours:F1} h" : $"{t.TotalMinutes:F0} min";
}
