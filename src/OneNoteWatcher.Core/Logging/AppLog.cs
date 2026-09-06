using System.Text;

namespace OneNoteWatcher.Core.Logging;

/// <summary>
/// Simple operational log: one file per day (<c>&lt;name&gt;-yyyy-MM-dd.log</c>) in the shared logs
/// folder, lines <c>HH:mm:ss.fff LEVEL message</c>. <see cref="Purge"/> deletes files older than the
/// retention window. Never throws — a logging failure must not take a detector down.
/// </summary>
public sealed class AppLog
{
    private readonly string _dir;
    private readonly string _name;
    private readonly object _gate = new();
    public string Directory => _dir;

    /// <summary>Writes that failed. Surfaced as an issue — losing the record is itself a problem.</summary>
    public int WriteFailures { get; private set; }

    /// <summary>Also echo to stdout — set for interactive runs (--replay, --audit) so the tool is not mute.</summary>
    public static bool EchoToConsole { get; set; }

    public AppLog(string dir, string name)
    {
        _dir = dir; _name = name;
        try { System.IO.Directory.CreateDirectory(dir); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    public string CurrentFile => Path.Combine(_dir, $"{_name}-{DateTime.Now:yyyy-MM-dd}.log");

    public void Info(string message) => Write("INFO ", message);
    public void Warn(string message) => Write("WARN ", message);
    public void Error(string message, Exception? ex = null) => Write("ERROR", ex is null ? message : $"{message}: {ex.GetType().Name}: {ex.Message}");

    private void Write(string level, string message)
    {
        var line = $"{DateTime.Now:HH:mm:ss.fff} {level} {message}{Environment.NewLine}";
        if (EchoToConsole) { try { Console.Write(line); } catch (IOException) { } }
        lock (_gate)
        {
            try { File.AppendAllText(CurrentFile, line, Encoding.UTF8); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { WriteFailures++; }
        }
    }

    /// <summary>Delete every *.log / *.txt in <paramref name="dir"/> last written more than <paramref name="retentionDays"/> ago.</summary>
    public static int Purge(string dir, int retentionDays)
    {
        if (retentionDays <= 0 || !System.IO.Directory.Exists(dir)) return 0;
        var cutoff = DateTime.UtcNow.AddDays(-retentionDays);
        var removed = 0;
        foreach (var f in System.IO.Directory.EnumerateFiles(dir))
        {
            if (!f.EndsWith(".log", StringComparison.OrdinalIgnoreCase) && !f.EndsWith(".txt", StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                if (File.GetLastWriteTimeUtc(f) < cutoff) { File.Delete(f); removed++; }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
        }
        return removed;
    }
}
