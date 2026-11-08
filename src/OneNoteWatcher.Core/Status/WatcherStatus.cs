using System.Text.Json;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Core.Status;

/// <summary>Last known sync state of one notebook.</summary>
public sealed record NotebookStatus(
    string Name, string? Gosid, DateTimeOffset? LastSyncUtc, bool? LastSuccess, string? LastCode, string? LastError,
    string? DisplayName = null)
{
    public string Label => DisplayName is null ? Name : $"{Name} (Display name: {DisplayName})";
}

/// <summary>
/// When the collector last saw OneNote complete HEALTHY sync activity for this section — a section
/// sync, a real-time session, or a page transfer that reported success.
///
/// This exists to answer one question the cloud check cannot answer on its own: <b>why did this
/// section's local timestamp move?</b> OneNote's search index re-stamps a section whenever it is
/// reconciled, including reconciliations that change nothing on the server, so "local newer than the
/// server copy" is not evidence of an unsent change. If the collector saw the section sync healthily at
/// that moment, the movement is explained and there is nothing stranded.
/// </summary>
public sealed record SectionSyncState(string Key, string Name, DateTimeOffset LastHealthySyncUtc);

/// <summary>
/// The collector's published view of the world, written to <c>status.json</c> in the shared dir and
/// read by the tray. It lets the tray show a meaningful icon whether or not OneNote is running:
/// "last known state" + timestamps rather than a blank grey.
/// </summary>
public sealed record WatcherStatus
{
    public DateTimeOffset UpdatedUtc { get; init; } = DateTimeOffset.UtcNow;
    public bool CollectorRunning { get; init; }
    public bool OneNoteRunning { get; init; }
    public bool? InternetAvailable { get; init; }
    public DateTimeOffset? LastSyncEventUtc { get; init; }
    /// <summary>When the collector last received ANY Office telemetry — the pipeline-liveness signal.</summary>
    public DateTimeOffset? LastTelemetryUtc { get; init; }
    public IReadOnlyList<NotebookStatus> Notebooks { get; init; } = [];
    /// <summary>Last healthy sync activity per section, keyed by <see cref="SectionKey"/>.</summary>
    public IReadOnlyList<SectionSyncState> Sections { get; init; } = [];
    public IReadOnlyList<SyncIssue> ActiveIssues { get; init; } = [];

    /// <summary>ETW records Windows dropped — a coverage gap the user should know about.</summary>
    public int EventsLost { get; init; }
    /// <summary>OneNote Storage payloads that could not be parsed (fail-closed counter).</summary>
    public int MalformedPayloads { get; init; }

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true };

    public void Save(string path)
    {
        try
        {
            var tmp = path + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(this, Json));
            File.Move(tmp, path, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { }
    }

    /// <summary>
    /// Read the published status, retrying briefly over the window in which <see cref="Save"/> replaces
    /// the file.
    ///
    /// Save writes a .tmp and moves it over this path, so for a few milliseconds every few seconds the
    /// path is mid-replace and an open fails. The tray polls far more often than that, so it loses the
    /// race regularly — and a null here is not a harmless miss: the caller drops EVERY issue the
    /// collector published and raises "the collector is not running — it has never reported" instead.
    /// Caught in the act on 2026-11-09 at 00:10:51, where the tray went 18 problems → 2 → 16 over six
    /// seconds while the collector was writing normally the whole time.
    ///
    /// Retries cover the replace window only; a genuinely absent, unreadable or corrupt file still
    /// returns null after them, so a collector that really has stopped is still reported.
    /// </summary>
    public static WatcherStatus? Load(string path)
    {
        for (var attempt = 0; ; attempt++)
        {
            try
            {
                if (!File.Exists(path)) return null;
                using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return JsonSerializer.Deserialize<WatcherStatus>(fs, Json);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
            {
                if (attempt >= 4) return null;
                Thread.Sleep(20 * (attempt + 1));
            }
        }
    }
}
