using System.Text;

namespace OneNoteWatcher.Core.History;

/// <summary>
/// Human-readable sync history, one file per day (<c>sync-history-yyyy-MM-dd.log</c>) in the shared
/// logs folder, one line per recorded sync result. Old days are removed by <see cref="Logging.AppLog.Purge"/>.
/// </summary>
public sealed class SyncHistoryLog
{
    private readonly string _dir;
    private readonly object _gate = new();

    public SyncHistoryLog(string dir)
    {
        _dir = dir;
        try { Directory.CreateDirectory(dir); } catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    public string Directory_ => _dir;
    public string CurrentFile => System.IO.Path.Combine(_dir, $"sync-history-{DateTime.Now:yyyy-MM-dd}.log");

    /// <summary>Appends that failed — losing a sync record is itself a problem, so it is surfaced.</summary>
    public int WriteFailures { get; private set; }

    /// <summary>Best-effort append; a write failure never takes the detector down, but it IS counted.</summary>
    public void Append(string line)
    {
        lock (_gate)
        {
            try { File.AppendAllText(CurrentFile, line + Environment.NewLine, Encoding.UTF8); }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { WriteFailures++; }
        }
    }
}
