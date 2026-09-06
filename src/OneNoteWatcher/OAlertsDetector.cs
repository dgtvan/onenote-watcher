using System.Diagnostics.Eventing.Reader;
using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher;

/// <summary>
/// Real-time, no-admin: Office writes every error dialog it shows to the <c>OAlerts</c> event log.
/// Subscribes for "Microsoft OneNote …" entries and turns error-looking ones into issues with a
/// category + fix from the dialog text. See docs/reference-data-sources.md §4.
/// </summary>
public sealed class OAlertsDetector : IDisposable
{
    public const string DetectorName = "oalerts";
    private readonly EventLogWatcher? _watcher;
    private readonly Action<SyncIssue> _onIssue;

    public OAlertsDetector(Action<SyncIssue> onIssue)
    {
        _onIssue = onIssue;
        try
        {
            var q = new EventLogQuery("OAlerts", PathType.LogName,
                "*[System[Provider[@Name='Microsoft Office 16 Alerts']]]");
            _watcher = new EventLogWatcher(q);
            _watcher.EventRecordWritten += OnRecord;
            _watcher.Enabled = true;
        }
        catch (EventLogException) { _watcher = null; }
    }

    public bool Available => _watcher is not null;

    private void OnRecord(object? sender, EventRecordWrittenEventArgs e)
    {
        string? text;
        try { text = e.EventRecord?.FormatDescription(); } catch { return; }
        if (text is null || !text.StartsWith("Microsoft OneNote", StringComparison.OrdinalIgnoreCase)) return;

        var body = text["Microsoft OneNote".Length..].Trim();
        var pIdx = body.IndexOf(" P1:", StringComparison.Ordinal);
        if (pIdx > 0) body = body[..pIdx].Trim();

        // confirmations are not errors
        if (body.StartsWith("Are you sure", StringComparison.OrdinalIgnoreCase)
            || body.Contains("Want to save", StringComparison.OrdinalIgnoreCase)) return;

        var ex = ErrorCatalog.Explain(0, body);
        var location = ExtractOneNoteUrl(body);
        _onIssue(new SyncIssue
        {
            Detector = DetectorName,
            Kind = IssueKind.ErrorDialog,
            NotebookName = location,
            Message = body.Length > 160 ? body[..160] + "…" : body,
            Category = ex.Category,
            Summary = $"OneNote showed an error dialog: \"{body}\"",
            Recommendation = ex.Category == FailureCategory.Unknown
                ? "Follow the dialog's instructions in OneNote; if it recurs, open File → Info → View Sync Status for the code."
                : ex.Recommendation,
            FirstSeen = DateTimeOffset.Now, LastSeen = DateTimeOffset.Now,
        });
    }

    private static string? ExtractOneNoteUrl(string s)
    {
        var i = s.IndexOf("onenote:", StringComparison.OrdinalIgnoreCase);
        if (i < 0) return null;
        var url = s[i..].Split(' ', '\r', '\n')[0];
        var seg = url.TrimEnd('/').Split('/');
        return seg.Length > 0 ? Uri.UnescapeDataString(seg[^1]) : url;
    }

    public void Dispose()
    {
        if (_watcher is null) return;
        _watcher.Enabled = false;
        _watcher.Dispose();
    }
}
