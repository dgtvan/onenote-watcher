using OneNoteWatcher.Core.Config;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Rules;

/// <summary>
/// The <c>[ignore]</c> mechanism (config.ini): suppress an issue by error code, message text,
/// notebook name, section name, originating detector, or OneNote event name. Suppressed issues are
/// still recorded in the history — they just never raise an alert.
///
/// The <c>events</c> key exists because the watcher is fail-closed: an event it cannot classify is
/// surfaced as a warning, and this is how you silence one once you are satisfied it is benign.
/// </summary>
public sealed class IgnoreRules
{
    private readonly IReadOnlyList<Pattern> _codes;
    private readonly IReadOnlyList<Pattern> _messages;
    private readonly IReadOnlyList<Pattern> _notebooks;
    private readonly IReadOnlyList<Pattern> _sections;
    private readonly IReadOnlyList<Pattern> _events;
    private readonly HashSet<string> _detectors;

    public IgnoreRules(
        IEnumerable<string> codes, IEnumerable<string> messages,
        IEnumerable<string> notebooks, IEnumerable<string> sections,
        IEnumerable<string> detectors, IEnumerable<string>? events = null)
    {
        _codes = codes.Select(Pattern.Parse).ToList();
        _messages = messages.Select(Pattern.Parse).ToList();
        _notebooks = notebooks.Select(Pattern.Parse).ToList();
        _sections = sections.Select(Pattern.Parse).ToList();
        _events = (events ?? []).Select(Pattern.Parse).ToList();
        _detectors = new HashSet<string>(detectors, StringComparer.OrdinalIgnoreCase);
    }

    public static IgnoreRules FromConfig(IniFile ini) => new(
        ini.GetList("ignore", "codes"),
        ini.GetList("ignore", "messages"),
        ini.GetList("ignore", "notebooks"),
        ini.GetList("ignore", "sections"),
        ini.GetList("ignore", "detectors"),
        ini.GetList("ignore", "events"));

    public static IgnoreRules Empty { get; } = new([], [], [], [], []);

    public bool IsIgnored(SyncIssue issue) =>
        _detectors.Contains(issue.Detector)
        || (issue.EventName is not null && _events.Any(p => p.IsMatch(issue.EventName)))
        || (issue.Code is not null && _codes.Any(p => p.IsMatch(issue.Code)))
        || (issue.Message.Length > 0 && _messages.Any(p => p.IsMatch(issue.Message)))
        || (issue.NotebookName is not null && _notebooks.Any(p => p.IsMatch(issue.NotebookName)))
        || (issue.SectionName is not null && _sections.Any(p => p.IsMatch(issue.SectionName)));
}
