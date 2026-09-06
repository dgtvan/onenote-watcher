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

    public static WatcherStatus? Load(string path)
    {
        try
        {
            if (!File.Exists(path)) return null;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return JsonSerializer.Deserialize<WatcherStatus>(fs, Json);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException) { return null; }
    }
}
