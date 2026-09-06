using System.Text.Json;

namespace OneNoteWatcher.Core.Index;

/// <summary>
/// Office's MRU service cache lists recently opened OneNote notebooks with their OneDrive
/// <c>ResourceId</c> and display <c>FileName</c>/<c>DocumentUrl</c>. It is the offline way to turn
/// the <c>NotebookId_ResourceId</c> in sync events into a notebook name.
/// Path: %LOCALAPPDATA%\Microsoft\Office\16.0\MruServiceCache\&lt;identity&gt;\OneNote\Documents_*
/// </summary>
public sealed class MruReader
{
    private readonly string _root;
    private Dictionary<string, string> _byResourceId = new(StringComparer.OrdinalIgnoreCase);
    private DateTime _loadedAt;

    public MruReader(string? mruRoot = null)
    {
        _root = mruRoot ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            @"Microsoft\Office\16.0\MruServiceCache");
    }

    public void RefreshIfStale(TimeSpan maxAge)
    {
        if (DateTime.UtcNow - _loadedAt < maxAge) return;
        Reload();
    }

    public void Reload()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (Directory.Exists(_root))
        {
            foreach (var file in Directory.EnumerateFiles(_root, "Documents_*", SearchOption.AllDirectories))
            {
                if (!file.Replace('\\', '/').Contains("/OneNote/", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    using var doc = JsonDocument.Parse(File.ReadAllText(file));
                    if (doc.RootElement.ValueKind != JsonValueKind.Array) continue;
                    foreach (var item in doc.RootElement.EnumerateArray())
                    {
                        var rid = item.TryGetProperty("ResourceId", out var r) ? r.GetString() : null;
                        var name = item.TryGetProperty("FileName", out var f) ? f.GetString() : null;
                        if (rid is null || name is null) continue;
                        map[rid] = name;
                    }
                }
                catch (Exception ex) when (ex is IOException or JsonException) { }
            }
        }
        _byResourceId = map;
        _loadedAt = DateTime.UtcNow;
    }

    public string? NotebookNameByResourceId(string? resourceId) =>
        resourceId is null ? null : _byResourceId.GetValueOrDefault(resourceId);
}
