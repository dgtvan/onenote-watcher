using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Index;
using System.Text.Json;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

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
    /// <summary>Per section: the server timestamp at the last poll where OneDrive was NOT behind the local
    /// copy. This is positive proof that the section's content reached the cloud — see
    /// <see cref="CloudConfirmedAfter"/>.</summary>
    private readonly Dictionary<string, DateTimeOffset> _cloudConfirmed = new(StringComparer.OrdinalIgnoreCase);

    private readonly TimeSpan _unconfirmedTtl;

    /// <param name="unconfirmedTtl">How long an UNCORROBORATED "not reaching OneDrive" claim may stand
    /// before it is abandoned — see the give-up rule in <see cref="Evaluate(IReadOnlyList{LocalSection},
    /// GraphSnapshot, bool, bool?, DateTimeOffset?, IReadOnlyList{SectionSyncState}?)"/>.</param>
    public OutcomeDetector(TimeSpan grace, TimeSpan? unconfirmedTtl = null)
    {
        _grace = grace;
        // not specified -> the two-hour default; explicitly zero or negative -> never give up
        _unconfirmedTtl = unconfirmedTtl switch
        {
            null => TimeSpan.FromHours(2),
            { } t when t <= TimeSpan.Zero => TimeSpan.MaxValue,
            { } t => t,
        };
    }

    /// <summary>Sections abandoned during the last <c>Evaluate</c>, with the reason — for the log.</summary>
    public IReadOnlyList<string> LastAbandoned { get; private set; } = [];

    public int BaselinedSections => _baseline.Count;

    public static IReadOnlyList<LocalSection> LocalSectionsFrom(SearchIndexReader index) =>
        index.Sections().Select(s => new LocalSection(index.NotebookNameOf(s), s.Title ?? "", index.NewestIn(s))).ToList();

    public IReadOnlyList<SyncIssue> Evaluate(SearchIndexReader index, GraphSnapshot server, bool oneNoteRunning, bool? internet, DateTimeOffset? now = null,
        IReadOnlyList<SectionSyncState>? collectorSections = null)
        => Evaluate(LocalSectionsFrom(index), server, oneNoteRunning, internet, now, collectorSections);

    /// <param name="collectorSections">
    /// Per-section full-sync successes observed by the ETW collector, when it is running. These are
    /// OneNote's own results and outrank anything inferred here from timestamps — see the note on
    /// <c>collectorSynced</c> below.
    /// </param>
    public IReadOnlyList<SyncIssue> Evaluate(IReadOnlyList<LocalSection> local, GraphSnapshot server, bool oneNoteRunning, bool? internet, DateTimeOffset? now = null,
        IReadOnlyList<SectionSyncState>? collectorSections = null)
    {
        var at = now ?? DateTimeOffset.UtcNow;
        var issues = new List<SyncIssue>();
        var abandoned = new List<string>();

        foreach (var gs in server.Sections)
        {
            // Prefer a notebook-qualified match. The name-only fallback exists because the local index
            // reports a notebook by the nickname set in OneNote ("Van") while Graph reports its real name
            // ("Note") — but it is taken ONLY when the section name is unique, since OneNote creates a
            // "Quick Notes" in every notebook and matching the wrong one compares two unrelated sections.
            var byName = local.Where(s => Eq(s.Section, gs.DisplayName)).ToList();
            var match = byName.FirstOrDefault(s => Eq(s.Notebook, gs.NotebookName))
                        ?? (byName.Count == 1 ? byName[0] : null);
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

            // The collector saw this section sync healthily at the moment its local timestamp moved, so
            // the movement is a sync re-stamp, not an unsent edit. That outranks both inputs this class
            // infers from, each of which is known to mislead: Graph's lastModifiedDateTime can lag, and
            // the local index re-stamps a section on ANY reconciliation — including one that changes
            // nothing on the server, which is precisely when the server timestamp will not move either.
            var collectorSynced = SectionSyncedAt(collectorSections, gs.Id, localNewest);

            // OneDrive is not behind this PC for this section: its content is safely on the server as of
            // serverNewest. Recorded so a real-time sync error the collector saw EARLIER than this can be
            // retired on direct evidence instead of standing forever (OneNote does not always emit a
            // success event after it recovers).
            if (serverCaughtUp)
            {
                if (!_cloudConfirmed.TryGetValue(key, out var had) || serverNewest > had)
                    _cloudConfirmed[key] = serverNewest;
            }

            if (localMoved && (serverCaughtUp || serverMoved || collectorSynced))
            {
                _baseline[key] = (localNewest, serverNewest); _aheadSince.Remove(key);   // synced
                continue;
            }
            if (localMoved && !serverCaughtUp && !collectorSynced)
            {
                if (!_aheadSince.TryGetValue(key, out var since)) { _aheadSince[key] = since = at; }

                // GIVE UP on a claim we have failed to corroborate.
                //
                // This comparison can be permanently wrong and never falsifiable. The local index is
                // re-stamped whenever a section is reconciled, including reconciliations that change
                // nothing on the server — and in exactly that case the server timestamp will never
                // advance either, so "local ahead" becomes a state with no way out. Measured 2026-09-06:
                // Quick Notes sat "unsynced" for 4.3 h while OneNote reported the notebook up to date.
                //
                // Abandoning it is only safe while we would have SEEN a genuine failure, so it requires
                // both: the collector fresh enough to be reporting (collectorSections non-null) and
                // OneNote running, i.e. it has had the chance to sync and to report an error and did
                // neither. With OneNote closed, or the collector down, the claim stands indefinitely —
                // which is the "edited, then shut down" case this check exists for.
                if (_unconfirmedTtl != TimeSpan.MaxValue
                    && collectorSections is not null && oneNoteRunning && at - since >= _unconfirmedTtl)
                {
                    _baseline[key] = (localNewest, serverNewest);
                    _aheadSince.Remove(key);
                    abandoned.Add($"{gs.NotebookName} / {gs.DisplayName}: no corroboration after " +
                                  $"{Human(at - since)} with OneNote running and the collector reporting — " +
                                  "treating the local timestamp as a sync re-stamp, not an unsent change");
                    continue;
                }

                if (at - since >= _grace)
                {
                    var ex = ErrorCatalog.UploadStuck(oneNoteRunning, internet);
                    issues.Add(new SyncIssue
                    {
                        Detector = DetectorName, Kind = IssueKind.UploadStuck,
                        NotebookName = match.Notebook ?? gs.NotebookName, SectionName = gs.DisplayName, SectionId = gs.Id,
                        Message = $"a local change at {localNewest.ToLocalTime():HH:mm} has not reached OneDrive after {Human(at - since)}",
                        Category = ex.Category, Summary = ex.Summary, Recommendation = ex.Recommendation,
                        TechnicalDetail = $"local={localNewest:u} server={serverNewest:u} graph-section={gs.Id}",
                        EvidenceUtc = localNewest,
                        FirstSeen = since, LastSeen = at,
                    });
                }
                continue;
            }
            if (!localMoved && serverMoved && oneNoteRunning && serverNewest - localNewest > _grace
                && !SectionSyncedAt(collectorSections, gs.Id, serverNewest))
            {
                _baseline[key] = (localNewest, serverNewest);
                issues.Add(new SyncIssue
                {
                    Detector = DetectorName, Kind = IssueKind.DownloadStuck,
                    NotebookName = match.Notebook ?? gs.NotebookName, SectionName = gs.DisplayName, SectionId = gs.Id,
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
        LastAbandoned = abandoned;
        return issues;
    }

    /// <summary>
    /// True when the collector saw healthy sync activity for this section at (or just before)
    /// <paramref name="at"/> — i.e. the local timestamp moving to that moment is explained by a sync
    /// rather than by an unsent edit. The small tolerance exists because the index stamp and the
    /// telemetry timestamp are written by different code paths a second or so apart.
    ///
    /// False whenever the collector is not running or the section is unmatched: an absent result is
    /// never taken as reassurance.
    /// </summary>
    private static bool SectionSyncedAt(IReadOnlyList<SectionSyncState>? sections, string? graphSectionId, DateTimeOffset at)
    {
        if (sections is null || sections.Count == 0) return false;
        var key = SectionKey.Normalize(graphSectionId);
        if (key is null) return false;
        var hit = sections.FirstOrDefault(x => x.Key == key);
        return hit is not null && hit.LastHealthySyncUtc >= at - SyncTolerance;
    }

    /// <summary>
    /// Did OneDrive confirm content for this section AFTER <paramref name="t"/>? True only when a poll
    /// actually compared the two sides and found the server at or ahead of this PC, and the server's own
    /// timestamp is later than <paramref name="t"/>. Anything less returns false, so a failure we cannot
    /// positively disprove keeps standing.
    /// </summary>
    public bool CloudConfirmedAfter(string? notebook, string? section, DateTimeOffset t)
    {
        if (string.IsNullOrEmpty(section)) return false;
        if (notebook is not null && _cloudConfirmed.TryGetValue($"{notebook}/{section}", out var exact))
            return exact > t;
        // notebook names can differ between the ETW event and Graph; fall back to a unique section match
        var hits = _cloudConfirmed.Where(kv => kv.Key.EndsWith("/" + section, StringComparison.OrdinalIgnoreCase)).ToList();
        return hits.Count == 1 && hits[0].Value > t;
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
