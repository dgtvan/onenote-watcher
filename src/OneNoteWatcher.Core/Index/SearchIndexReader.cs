using Microsoft.Data.Sqlite;

namespace OneNoteWatcher.Core.Index;

/// <summary>An entity from OneNote's FullTextSearchIndex (docs/reference-data-sources.md §1).</summary>
public sealed record IndexEntity(
    int Type, string Goid, string? Gosid, string? ParentGoid, string? Title,
    DateTimeOffset? LastModifiedUtc, string NotebookGoidPrefix)
{
    public bool IsNotebook => Type == 4;
    public bool IsSectionGroup => Type == 3;
    public bool IsSection => Type == 2;
    public bool IsPage => Type == 1;
}

/// <summary>
/// Read-only view of the local search index: resolves notebook GOSIDs and page/section GOIDs to
/// names and gives per-section newest LastModifiedTime. Never touches OneNote — plain SQLite reads
/// with <c>mode=ro</c>; tolerant of locks (returns the last good snapshot).
/// </summary>
public sealed class SearchIndexReader
{
    private readonly string _dir;
    private Snapshot _snap = Snapshot.Empty;
    private DateTime _loadedAt;

    public SearchIndexReader(string? indexDir = null)
    {
        _dir = indexDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Microsoft\OneNote\16.0\FullTextSearchIndex");
    }

    public bool Available => Directory.Exists(_dir);

    /// <summary>Why the last load was incomplete, or null when the index read cleanly. Surfaced as an issue.</summary>
    public string? DegradedReason { get; private set; }

    /// <summary>Reload if older than <paramref name="maxAge"/>. Safe to call often.</summary>
    public void RefreshIfStale(TimeSpan maxAge)
    {
        if (DateTime.UtcNow - _loadedAt < maxAge) return;
        Reload();
    }

    public void Reload()
    {
        var byGosid = new Dictionary<string, IndexEntity>(StringComparer.OrdinalIgnoreCase);
        var byGoid = new Dictionary<string, IndexEntity>(StringComparer.OrdinalIgnoreCase);
        if (!Directory.Exists(_dir))
        {
            _snap = Snapshot.Empty; _loadedAt = DateTime.UtcNow;
            DegradedReason = $"folder not found: {_dir}";
            return;
        }

        var failed = 0; var total = 0; string? firstError = null;
        foreach (var db in Directory.EnumerateFiles(_dir, "*.db"))
        {
            total++;
            try
            {
                using var con = new SqliteConnection($"Data Source={db};Mode=ReadOnly");
                con.Open();
                using var cmd = con.CreateCommand();
                cmd.CommandText = "SELECT Type, GOID, GOSID, ParentGOID, Title, LastModifiedTime FROM Entities";
                using var r = cmd.ExecuteReader();
                var nbPrefix = Path.GetFileNameWithoutExtension(db);
                while (r.Read())
                {
                    var e = new IndexEntity(
                        r.GetInt32(0), r.GetString(1),
                        r.IsDBNull(2) ? null : r.GetString(2),
                        r.IsDBNull(3) ? null : r.GetString(3),
                        r.IsDBNull(4) ? null : r.GetString(4),
                        r.IsDBNull(5) ? null : FileTime(r.GetInt64(5)),
                        nbPrefix);
                    byGoid[e.Goid] = e;
                    if (e.Gosid is not null) byGosid[e.Gosid] = e;
                }
            }
            catch (Exception ex) when (ex is SqliteException or IOException or UnauthorizedAccessException)
            {
                // a file we could not read means our view of the notebooks is incomplete — never silent
                failed++; firstError ??= $"{Path.GetFileName(db)}: {ex.Message}";
            }
        }
        _snap = new Snapshot(byGosid, byGoid);
        _loadedAt = DateTime.UtcNow;
        DegradedReason = failed == 0
            ? (total == 0 ? "no index files found" : null)
            : $"{failed} of {total} index file(s) unreadable — {firstError}";
    }

    /// <summary>Notebook/section/page by stable id; accepts ids with or without the trailing {B0}.</summary>
    public IndexEntity? ByGosid(string? gosid)
    {
        if (gosid is null) return null;
        var key = gosid.EndsWith("{B0}", StringComparison.Ordinal) ? gosid[..^4] : gosid;
        return _snap.ByGosid.GetValueOrDefault(key);
    }

    public IndexEntity? ByGoid(string? goid) => goid is null ? null : _snap.ByGoid.GetValueOrDefault(goid);

    /// <summary>The Type=4 root entity of a page/section by walking ParentGOID.</summary>
    public IndexEntity? NotebookOf(IndexEntity? e)
    {
        var cur = e; var hops = 0;
        while (cur is not null && !cur.IsNotebook && hops++ < 10)
            cur = cur.ParentGoid is null ? null : _snap.ByGoid.GetValueOrDefault(cur.ParentGoid);
        return cur?.IsNotebook == true ? cur : null;
    }

    /// <summary>Notebook display name for a page/section entity.</summary>
    public string? NotebookNameOf(IndexEntity? e) => NotebookOf(e)?.Title;

    /// <summary>Section display name for a page entity (its Type=2 ancestor).</summary>
    public string? SectionNameOf(IndexEntity? e)
    {
        var cur = e; var hops = 0;
        while (cur is not null && !cur.IsSection && hops++ < 10)
            cur = cur.ParentGoid is null ? null : _snap.ByGoid.GetValueOrDefault(cur.ParentGoid);
        return cur?.IsSection == true ? cur.Title : null;
    }

    public IEnumerable<IndexEntity> Notebooks() => _snap.ByGoid.Values.Where(e => e.IsNotebook);
    public IEnumerable<IndexEntity> Sections() => _snap.ByGoid.Values.Where(e => e.IsSection);
    public IEnumerable<IndexEntity> All() => _snap.ByGoid.Values;

    /// <summary>Newest LastModifiedTime in a subtree (section or notebook), per docs: use MAX over children.</summary>
    public DateTimeOffset? NewestIn(IndexEntity root)
    {
        DateTimeOffset? best = root.LastModifiedUtc;
        foreach (var e in _snap.ByGoid.Values)
        {
            if (!IsDescendant(e, root.Goid)) continue;
            if (e.LastModifiedUtc is { } t && (best is null || t > best)) best = t;
        }
        return best;
    }

    private bool IsDescendant(IndexEntity e, string ancestorGoid)
    {
        var cur = e; var hops = 0;
        while (cur?.ParentGoid is not null && hops++ < 10)
        {
            if (cur.ParentGoid == ancestorGoid) return true;
            cur = _snap.ByGoid.GetValueOrDefault(cur.ParentGoid);
        }
        return false;
    }

    private static DateTimeOffset? FileTime(long ft) =>
        ft <= 0 ? null : DateTimeOffset.FromFileTime(ft).ToUniversalTime();

    private sealed record Snapshot(
        Dictionary<string, IndexEntity> ByGosid, Dictionary<string, IndexEntity> ByGoid)
    {
        public static readonly Snapshot Empty = new(new(), new());
    }
}
