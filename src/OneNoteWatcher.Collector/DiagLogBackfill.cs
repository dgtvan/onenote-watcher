using System.Text;

namespace OneNoteWatcher.Collector;

/// <summary>
/// Post-session backfill: reads Office diagnostic logs of *finished* OneNote sessions (the live one is
/// locked deny-read) and feeds their Storage SendEvent rows through the same detector, so history is
/// complete even for periods when the collector was not running. Processed session GUIDs are
/// remembered so nothing is double-logged. See docs/reference-data-sources.md §5.
/// </summary>
public sealed class DiagLogBackfill
{
    private readonly string _diagDir;
    private readonly string _stateFile;
    private readonly TimeSpan _maxAge;

    public DiagLogBackfill(string sharedDir, string? diagDir = null, TimeSpan? maxAge = null)
    {
        _diagDir = diagDir ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Temp\Diagnostics\ONENOTE");
        _stateFile = Path.Combine(sharedDir, "processed-sessions.txt");
        _maxAge = maxAge ?? TimeSpan.FromDays(7);
    }

    /// <summary>Returns the number of session files ingested.</summary>
    public int Run(EtwSyncDetector detector, OneNoteWatcher.Core.Logging.AppLog? log = null)
    {
        if (!Directory.Exists(_diagDir)) return 0;
        var done = File.Exists(_stateFile)
            ? new HashSet<string>(File.ReadAllLines(_stateFile), StringComparer.OrdinalIgnoreCase)
            : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var ingested = 0;
        var files = Directory.EnumerateFiles(_diagDir, "Primary*.log")
            .Select(f => new FileInfo(f))
            .Where(f => DateTime.UtcNow - f.LastWriteTimeUtc < _maxAge)
            .OrderBy(f => f.LastWriteTimeUtc);

        foreach (var fi in files)
        {
            var key = fi.Name;
            if (done.Contains(key)) continue;

            string text;
            try
            {
                using var fs = new FileStream(fi.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var sr = new StreamReader(fs, Encoding.UTF8);
                text = sr.ReadToEnd().Replace("\0", "");
            }
            catch (IOException) { continue; } // still locked = live session; try next time

            var rows = new List<(DateTime t, string msg)>();
            foreach (var line in text.Split('\n'))
            {
                var i = line.IndexOf("SendEvent {", StringComparison.Ordinal);
                if (i < 0 || !line.Contains("\"Office.OneNote.Storage.", StringComparison.Ordinal)) continue;
                var msg = line[i..].TrimEnd('\r', '\t');
                rows.Add((DateTime.UtcNow, msg));
            }
            foreach (var (t, msg) in rows)
                detector.OnMessage(new OfficeLogMessage(t, OfficeEtw.TelemetryCategory, msg));

            done.Add(key); ingested++;
            log?.Info($"backfill {fi.Name}: {rows.Count} storage events");
        }

        try { File.WriteAllLines(_stateFile, done); } catch (IOException) { }
        return ingested;
    }
}
