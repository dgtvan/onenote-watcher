using OneNoteWatcher.Core.Config;

namespace OneNoteWatcher.Core.Rules;

/// <summary>
/// The <c>[transient]</c> config section — the ONLY place a genuine failure is deliberately held back,
/// and only for error codes you list yourself.
///
/// The watcher has two states, error and success, so by default every failure OneNote reports raises an
/// error immediately. Some codes are genuinely retryable (a server timeout that succeeds a second
/// later), and you may not want to look at those unless they persist. Listing such a code here says:
/// "only tell me if this happens <c>threshold</c> times within <c>window_minutes</c>". Held-back
/// occurrences are still written to the sync history, so nothing is lost — only the red icon is delayed.
/// An empty <c>codes</c> list (the default) means nothing is ever held back.
/// </summary>
public sealed class TransientPolicy
{
    private readonly IReadOnlyList<Pattern> _codes;
    private readonly int _threshold;
    private readonly TimeSpan _window;
    private readonly Dictionary<string, List<DateTimeOffset>> _seen = new(StringComparer.OrdinalIgnoreCase);

    public TransientPolicy(IEnumerable<string> codes, int threshold, int windowMinutes)
    {
        _codes = codes.Select(Pattern.Parse).ToList();
        _threshold = Math.Max(1, threshold);
        _window = TimeSpan.FromMinutes(Math.Max(1, windowMinutes));
    }

    public static TransientPolicy FromConfig(IniFile ini) => new(
        ini.GetList("transient", "codes"),
        ini.GetInt("transient", "threshold", 3),
        ini.GetInt("transient", "window_minutes", 30));

    /// <summary>Nothing is ever held back.</summary>
    public static TransientPolicy Default { get; } = new([], 3, 30);

    /// <summary>
    /// True when this occurrence should be held back: the code is one the user listed AND it has not yet
    /// reached the threshold inside the window. Records the occurrence either way.
    /// </summary>
    public bool ShouldHoldBack(string scopeKey, string? code, DateTimeOffset at)
    {
        if (code is null || !_codes.Any(p => p.IsMatch(code))) return false;   // not opted in → raise now

        var key = scopeKey + "|" + code;
        if (!_seen.TryGetValue(key, out var times)) _seen[key] = times = [];
        times.Add(at);
        times.RemoveAll(t => at - t > _window);
        return times.Count < _threshold;
    }

    /// <summary>Reset the counter for a scope after it recovers.</summary>
    public void Forget(string scopeKey)
    {
        foreach (var k in _seen.Keys.Where(k => k.StartsWith(scopeKey + "|", StringComparison.OrdinalIgnoreCase)).ToList())
            _seen.Remove(k);
    }
}
