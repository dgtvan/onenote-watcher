namespace OneNoteWatcher.Core.Config;

/// <summary>
/// Minimal INI reader: <c>[section]</c> headers, <c>key = value</c> lines, <c>;</c> or <c>#</c>
/// comments (whole-line or trailing). No dependencies. Keys are case-insensitive.
/// </summary>
public sealed class IniFile
{
    private readonly Dictionary<string, Dictionary<string, string>> _sections =
        new(StringComparer.OrdinalIgnoreCase);

    public static IniFile Parse(string text)
    {
        var ini = new IniFile();
        var current = "";
        ini._sections[current] = new(StringComparer.OrdinalIgnoreCase);

        foreach (var rawLine in text.Split('\n'))
        {
            var line = rawLine.Trim();
            if (line.Length == 0 || line[0] == ';' || line[0] == '#') continue;

            if (line[0] == '[')
            {
                var end = line.IndexOf(']');
                if (end > 1)
                {
                    current = line[1..end].Trim();
                    if (!ini._sections.ContainsKey(current))
                        ini._sections[current] = new(StringComparer.OrdinalIgnoreCase);
                }
                continue;
            }

            var eq = line.IndexOf('=');
            if (eq < 0) continue;
            var key = line[..eq].Trim();
            var value = StripInlineComment(line[(eq + 1)..]).Trim();
            ini._sections[current][key] = value;
        }
        return ini;
    }

    public static IniFile Load(string path) => Parse(File.ReadAllText(path));

    // A ';' or '#' preceded by whitespace starts a trailing comment. (A '#' inside a value with no
    // leading space — rare here — is kept.)
    private static string StripInlineComment(string value)
    {
        for (var i = 1; i < value.Length; i++)
            if ((value[i] == ';' || value[i] == '#') && char.IsWhiteSpace(value[i - 1]))
                return value[..i];
        return value;
    }

    public string? Get(string section, string key) =>
        _sections.TryGetValue(section, out var s) && s.TryGetValue(key, out var v) ? v : null;

    public string Get(string section, string key, string fallback) =>
        Get(section, key) is { Length: > 0 } v ? v : fallback;

    public int GetInt(string section, string key, int fallback) =>
        int.TryParse(Get(section, key), out var v) ? v : fallback;

    public bool GetBool(string section, string key, bool fallback) =>
        Get(section, key) is { } v ? v.Equals("true", StringComparison.OrdinalIgnoreCase) || v == "1" : fallback;

    /// <summary>A comma-separated value split into trimmed, non-empty items.</summary>
    public IReadOnlyList<string> GetList(string section, string key) =>
        (Get(section, key) ?? "")
        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
}
