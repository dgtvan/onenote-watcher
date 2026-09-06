using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Index;

/// <summary>
/// Resolved human names for a sync event's location. <see cref="Notebook"/> is the notebook's original
/// name as OneDrive/Graph/OneNote's file list show it ("Note"); <see cref="NotebookDisplayName"/> is the
/// nickname the user set in OneNote's navigation pane ("Van") when it differs, else null.
/// </summary>
public sealed record ResolvedNames(string? Notebook, string? Section, string? Page, string? NotebookDisplayName = null)
{
    /// <summary>"Note (Display name: Van)" or just "Note".</summary>
    public string? NotebookLabel => Notebook is null ? NotebookDisplayName
        : NotebookDisplayName is null ? Notebook : $"{Notebook} (Display name: {NotebookDisplayName})";
}

/// <summary>
/// Turns the ids inside a <see cref="SyncEvent"/> into notebook / section / page names, using only
/// local, read-only sources: the search index (notebook GOSID, page GOID), the Office MRU cache
/// (notebook OneDrive resource id), and a learned rid→GOSID map (NotebookSyncResult events carry
/// both, so section events that only carry the rid still get the index's display name).
/// Section names by OneDrive resource id come from <see cref="SectionNameByResourceId"/> (Graph map).
/// </summary>
public sealed class NameResolver
{
    private readonly SearchIndexReader _index;
    private readonly MruReader _mru;
    private readonly Dictionary<string, string> _ridToGosid = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _gosidToRid = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Optional external lookup (Graph-derived map): (sectionResourceId, sectionGosid) → display name.</summary>
    public Func<string?, string?, string?>? SectionNameByResourceId { get; set; }

    public NameResolver(SearchIndexReader? index = null, MruReader? mru = null)
    {
        _index = index ?? new SearchIndexReader();
        _mru = mru ?? new MruReader();
    }

    public void Refresh(TimeSpan maxAge)
    {
        _index.RefreshIfStale(maxAge);
        _mru.RefreshIfStale(maxAge);
    }

    /// <summary>Non-null when the local index could not be fully read — surfaced as a health issue.</summary>
    public string? IndexDegradedReason => _index.DegradedReason;

    /// <summary>Remember rid→GOSID pairs seen on notebook events.</summary>
    public void Learn(SyncEvent e)
    {
        if (e.NotebookResourceId is not null && e.NotebookGosid is not null)
        {
            _ridToGosid[e.NotebookResourceId] = e.NotebookGosid;
            _gosidToRid[e.NotebookGosid] = e.NotebookResourceId;
        }
    }

    public ResolvedNames Resolve(SyncEvent e)
    {
        Learn(e);
        var gosid = e.NotebookGosid
                    ?? (e.NotebookResourceId is not null ? _ridToGosid.GetValueOrDefault(e.NotebookResourceId) : null);
        var rid = e.NotebookResourceId
                  ?? (e.NotebookGosid is not null ? _gosidToRid.GetValueOrDefault(e.NotebookGosid) : null);

        // original name = MRU/Graph/OneDrive ("Note"); display name = index root Title, the nickname the
        // user set in OneNote's navigation pane ("Van"). Show both when they differ.
        string? original = _mru.NotebookNameByResourceId(rid);
        string? display = _index.ByGosid(gosid)?.Title;
        string? section = null;
        string? page = null;

        if (e.PageGoid is not null)
        {
            var pg = _index.ByGoid(e.PageGoid);
            if (pg is not null)
            {
                page = pg.Title;
                section ??= _index.SectionNameOf(pg);
                var root = _index.NotebookOf(pg);
                display ??= root?.Title;
                if (original is null && root?.Gosid is not null)
                    original = _mru.NotebookNameByResourceId(_gosidToRid.GetValueOrDefault(root.Gosid));
            }
        }
        string? notebook = original ?? display;
        string? notebookDisplay = display is not null && !string.Equals(display, notebook, StringComparison.OrdinalIgnoreCase) ? display : null;

        if (section is null && (e.SectionResourceId is not null || e.SectionGosid is not null) && SectionNameByResourceId is not null)
        {
            var full = SectionNameByResourceId(e.SectionResourceId, e.SectionGosid);
            if (full is not null)
            {
                // map values are "Notebook / Section"; keep just the section part when notebook is known
                var slash = full.IndexOf(" / ", StringComparison.Ordinal);
                section = slash > 0 ? full[(slash + 3)..] : full;
                notebook ??= slash > 0 ? full[..slash] : null;
            }
        }

        return new ResolvedNames(notebook, section, page, notebookDisplay);
    }
}
