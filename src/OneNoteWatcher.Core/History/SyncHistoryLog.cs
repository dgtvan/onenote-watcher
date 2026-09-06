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

    /// <summary>
    /// Best-effort append; a write failure never takes the detector down, but it IS counted.
    ///
    /// Two processes write this file — the collector records syncs, the tray records a recovery the
    /// collector cannot see (it holds no Graph token). So the handle is opened FileShare.ReadWrite and
    /// each line is written in a single call, and a brief collision is retried rather than counted as a
    /// failure: a spurious WriteFailures bump would raise a "logging is broken" alert of its own.
    /// </summary>
    public void Append(string line)
    {
        lock (_gate)
        {
            var payload = Encoding.UTF8.GetBytes(line + Environment.NewLine);
            for (var attempt = 0; ; attempt++)
            {
                try
                {
                    using var fs = new FileStream(CurrentFile, FileMode.Append, FileAccess.Write,
                        FileShare.ReadWrite, bufferSize: 0, FileOptions.WriteThrough);
                    fs.Write(payload, 0, payload.Length);
                    return;
                }
                catch (IOException) when (attempt < 4) { Thread.Sleep(10 * (attempt + 1)); }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    WriteFailures++;
                    return;
                }
            }
        }
    }
}
