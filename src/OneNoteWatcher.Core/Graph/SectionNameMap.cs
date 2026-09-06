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

    /// <summary>Candidate keys: the Graph id and its hex tokens, the section-id GUID, the .one file name.</summary>
    internal static IEnumerable<string> KeysOf(GraphSection s)
    {
        yield return s.Id;
        foreach (Match t in Regex.Matches(s.Id, "[0-9a-fA-F]{16,}")) yield return t.Value;
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
            var tok = resourceId.Split('!')[^1].TrimStart('s', 'S');
            if (tok.Length >= 16)
                foreach (var kv in _byKey)
                    if (kv.Key.Contains(tok, StringComparison.OrdinalIgnoreCase)) return kv.Value;
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
