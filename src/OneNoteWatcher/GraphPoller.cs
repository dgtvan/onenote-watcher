using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Logging;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher;

/// <summary>
/// Runs in the tray (user context): fetches Graph, publishes the section-name map for the collector,
/// and evaluates the stateful no-admin <see cref="OutcomeDetector"/> baseline. Never prompts on its own.
/// </summary>
public sealed class GraphPoller
{
    public const string SignInDetector = "graph";
    private readonly GraphAuth _auth;
    private readonly GraphClient _client;
    private readonly SearchIndexReader _index = new();
    private readonly OutcomeDetector _outcome;
    private readonly string _sectionMapPath;
    private readonly string _statePath;
    private readonly AppLog? _log;

    public GraphSnapshot? LastSnapshot { get; private set; }
    public DateTimeOffset? LastFetchUtc { get; private set; }
    /// <summary>Awake-time stamp of the last successful fetch, so a sleeping machine is not mistaken for
    /// a stalled poller (see <see cref="AwakeClock"/>).</summary>
    public TimeSpan LastFetchAwake { get; private set; } = AwakeClock.Stamp();

    /// <summary>Running time since the last successful fetch, excluding time the machine was asleep.</summary>
    public TimeSpan SinceLastFetch =>
        LastFetchUtc is { } f ? AwakeClock.Elapsed(f, DateTimeOffset.UtcNow, LastFetchAwake) : TimeSpan.MaxValue;
    public string? LastError { get; private set; }
    public bool SignedIn { get; private set; }

    /// <summary>
    /// True when OneDrive has positively confirmed this section's content at a time later than
    /// <paramref name="t"/>. Used to retire a real-time sync error that OneNote never followed with a
    /// success event. Requires a recent successful poll — a stale or failed cloud check confirms nothing.
    /// </summary>
    public bool CloudConfirmedAfter(string? notebook, string? section, DateTimeOffset t) =>
        SignedIn && LastError is null
        && SinceLastFetch < TimeSpan.FromMinutes(30)
        && _outcome.CloudConfirmedAfter(notebook, section, t);

    public GraphPoller(GraphAuth auth, TimeSpan grace, string sharedDir, AppLog? log = null, TimeSpan? unconfirmedTtl = null)
    {
        _auth = auth; _log = log;
        _client = new GraphClient(async ct => await _auth.TryGetTokenSilentAsync(ct) ?? throw new InvalidOperationException("not signed in"));
        _outcome = new OutcomeDetector(grace, unconfirmedTtl);
        _sectionMapPath = Path.Combine(sharedDir, "section-names.json");
        _statePath = Path.Combine(sharedDir, "outcome-state.json");
        var restored = _outcome.LoadState(_statePath);
        if (restored > 0) _log?.Info($"restored cloud-check baseline for {restored} section(s) — a change stranded before a reboot is still flagged");
    }

    /// <param name="collectorSections">The collector's per-section sync results, when it is running.
    /// OneNote's own outcome beats this class's timestamp inference.</param>
    public async Task<IReadOnlyList<SyncIssue>> PollAsync(bool oneNoteRunning, bool? internet, CancellationToken ct,
        IReadOnlyList<SectionSyncState>? collectorSections = null)
    {
        var wasSignedIn = SignedIn;
        SignedIn = await _auth.TryGetTokenSilentAsync(ct) is not null;
        if (wasSignedIn != SignedIn) _log?.Info($"Graph signed in: {SignedIn}");
        if (!SignedIn) return [];   // surfaced by HealthIssues.GraphBlind (tray)
        try
        {
            LastSnapshot = await _client.FetchAsync(ct);
            LastFetchUtc = DateTimeOffset.UtcNow; LastFetchAwake = AwakeClock.Stamp(); LastError = null;
            var map = SectionNameMap.FromSnapshot(LastSnapshot);
            map.Save(_sectionMapPath);
            _index.RefreshIfStale(TimeSpan.FromMinutes(1));
            var issues = _outcome.Evaluate(_index, LastSnapshot, oneNoteRunning, internet, null, collectorSections);
            _outcome.SaveState(_statePath);   // survives tray restart / reboot
            _log?.Info($"graph poll: {LastSnapshot.Notebooks.Count} notebooks, {LastSnapshot.Sections.Count} sections, map keys={map.Count}, baselined={_outcome.BaselinedSections}, collector sections={collectorSections?.Count ?? 0}, outcome issues={issues.Count}");
            foreach (var i in issues) _log?.Warn($"outcome issue {i.Location}: {i.Message}");
            foreach (var a in _outcome.LastAbandoned) _log?.Info($"cloud check gave up on {a}");
            return issues;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or InvalidOperationException)
        {
            LastError = ex.Message;
            _log?.Error("graph fetch failed", ex);
            return [];
        }
    }
}
