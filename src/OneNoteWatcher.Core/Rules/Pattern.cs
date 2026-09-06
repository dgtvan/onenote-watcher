using System.Text.RegularExpressions;

namespace OneNoteWatcher.Core.Rules;

/// <summary>
/// A single match pattern from config: a glob (<c>*</c>, <c>?</c>, case-insensitive) or, when
/// wrapped in slashes, a regex — <c>/foo.*/</c> or <c>/foo.*/i</c>.
/// </summary>
public sealed class Pattern
{
    private readonly Regex _regex;
    public string Raw { get; }

    private Pattern(string raw, Regex regex) { Raw = raw; _regex = regex; }

    public bool IsMatch(string? value) => value is not null && _regex.IsMatch(value);

    public static Pattern Parse(string raw)
    {
        raw = raw.Trim();
        // /regex/ or /regex/i
        if (raw.Length >= 2 && raw[0] == '/')
        {
            var end = raw.LastIndexOf('/');
            if (end > 0)
            {
                var body = raw[1..end];
                var flags = raw[(end + 1)..];
                var opts = RegexOptions.CultureInvariant;
                if (flags.Contains('i', StringComparison.OrdinalIgnoreCase))
                    opts |= RegexOptions.IgnoreCase;
                return new Pattern(raw, new Regex(body, opts));
            }
        }
        // glob → regex, case-insensitive, anchored
        var pat = "^" + Regex.Escape(raw).Replace("\\*", ".*").Replace("\\?", ".") + "$";
        return new Pattern(raw, new Regex(pat, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant));
    }
}
