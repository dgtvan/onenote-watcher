using System.Text.Json;
using System.Text.RegularExpressions;

namespace OneNoteWatcher.Core.Graph;

/// <summary>
/// Section display names keyed by every identifier Graph exposes for a section, so the OneDrive
/// resource id (<c>…!s8d49…</c>) and the <c>section-id={GUID}</c> seen in sync events can be turned
/// into "ASW". Written by the tray (which holds the Graph token) to <c>section-names.json</c> in the
/// shared dir and read by the elevated collector, which has no user token.
/// </summary>
public sealed class SectionNameMap
{
    private readonly Dictionary<string, string> _byKey = new(StringComparer.OrdinalIgnoreCase);
    public int Count => _byKey.Count;

    public static SectionNameMap FromSnapshot(GraphSnapshot snap)
    {
        var m = new SectionNameMap();
        foreach (var s in snap.Sections)
        {
            var name = string.IsNullOrEmpty(s.NotebookName) ? s.DisplayName : $"{s.NotebookName} / {s.DisplayName}";
            foreach (var k in KeysOf(s)) m._byKey[k] = name;
        }
        return m;
    }

    /// <summary>Prefix for the canonical <see cref="SectionKey"/> entry, so it cannot collide with a raw id.</summary>
    private const string CanonPrefix = "key:";

    /// <summary>Candidate keys: the Graph id, its canonical key, its hex tokens, the section-id GUID, the .one file name.</summary>
    internal static IEnumerable<string> KeysOf(GraphSection s)
    {
        yield return s.Id;
        if (Model.SectionKey.Normalize(s.Id) is { } canon) yield return CanonPrefix + canon;
        foreach (Match t in Regex.Matches(s.Id, "[0-9a-fA-F]{16,}"))
        {
            // Never key on the DRIVE id — the token immediately before "!". It is identical for every
            // section in the drive, so all of them write that one entry and the last one wins; a lookup
            // then returns a confidently wrong section name. On this machine 16 sections shared a single
            // drive id, so the collision was the normal case, not an edge one.
            var isDriveId = t.Index + t.Length < s.Id.Length && s.Id[t.Index + t.Length] == '!';
            if (!isDriveId) yield return t.Value;
        }
        var g = Regex.Match(s.ClientUrl ?? "", @"section-id=\{?([0-9a-fA-F-]{36})\}?", RegexOptions.IgnoreCase);
        if (g.Success) yield return "{" + g.Groups[1].Value.ToUpperInvariant() + "}";
        if (s.WebUrl is not null)
        {
            var last = Uri.UnescapeDataString(s.WebUrl.Split('?')[0].TrimEnd('/').Split('/')[^1]);
            if (last.Length > 0) yield return "file:" + last;
        }
    }

    /// <summary>Resolve from a sync event's section OneDrive resource id and/or UnmappedGosid.</summary>
    public string? Lookup(string? resourceId, string? sectionGosid)
    {
        if (resourceId is not null)
        {
            if (_byKey.TryGetValue(resourceId, out var n)) return n;
            // canonical match: one key for every spelling of the same section, so a sync event's
            // "36B934175DC7E3A4!1242" finds the name stored under Graph's "0-36B934175DC7E3A4!1242"
            if (Model.SectionKey.Normalize(resourceId) is { } canon
                && _byKey.TryGetValue(CanonPrefix + canon, out var nc)) return nc;
            // Substring fallback, for a spelling neither exact nor canonical matching caught. Two guards,
            // both learned from the same failure: the token must be the ITEM part (after "!"), because a
            // bare drive id is a substring of every section id in that drive; and it must identify
            // exactly ONE section, or this returns whichever entry the dictionary happened to yield
            // first — a confidently wrong name, which is worse than no name at all.
            var bang = resourceId.LastIndexOf('!');
            if (bang >= 0)
            {
                var tok = resourceId[(bang + 1)..].TrimStart('s', 'S');
                if (tok.Length >= 16)
                {
                    var hits = _byKey.Where(kv => kv.Key.Contains(tok, StringComparison.OrdinalIgnoreCase))
                                     .Select(kv => kv.Value).Distinct(StringComparer.Ordinal).Take(2).ToList();
                    if (hits.Count == 1) return hits[0];
                }
            }
        }
        if (sectionGosid is not null)
        {
            var guid = Regex.Match(sectionGosid, "[0-9a-fA-F-]{36}");
            if (guid.Success && _byKey.TryGetValue("{" + guid.Value.ToUpperInvariant() + "}", out var n2)) return n2;
        }
        return null;
    }

    public void Save(string path)
    {
        try
        {
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(_byKey));
            File.Move(tmp, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    public static SectionNameMap Load(string path)
    {
        var m = new SectionNameMap();
        try
        {
            if (!File.Exists(path)) return m;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            foreach (var kv in JsonSerializer.Deserialize<Dictionary<string, string>>(fs) ?? new())
                m._byKey[kv.Key] = kv.Value;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { }
        return m;
    }
}
