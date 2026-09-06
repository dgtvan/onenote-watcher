using System.Text;
using OneNoteWatcher.Core.Config;
using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Index;

namespace OneNoteWatcher;

/// <summary>
/// `OneNoteWatcher.exe --graph-check`: one-shot report comparing every Graph section's server
/// lastModifiedDateTime with the local search index, plus what the OutcomeDetector would raise.
/// Writes %LOCALAPPDATA%\OneNoteWatcher\graph-check.txt (and prints the path). Diagnoses false alarms.
/// </summary>
public static class GraphCheck
{
    public static async Task<string> RunAsync(string configPath, string? outDir = null)
    {
        var ini = File.Exists(configPath) ? IniFile.Load(configPath) : IniFile.Parse("");
        outDir ??= Path.Combine(ini.Get("etw", "shared_dir", @"C:\ProgramData\OneNoteWatcher"), "logs");
        try { Directory.CreateDirectory(outDir); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            outDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OneNoteWatcher");
            Directory.CreateDirectory(outDir);
        }
        var outPath = Path.Combine(outDir, $"graph-check-{DateTime.Now:yyyy-MM-dd-HHmmss}.txt");
        var sb = new StringBuilder();
        sb.AppendLine($"OneNote Sync Watcher — cloud check report  {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}");

        var clientId = ini.Get("graph", "client_id");
        if (string.IsNullOrWhiteSpace(clientId)) { sb.AppendLine("no [graph] client_id in config.ini"); return Write(outPath, sb); }
        var auth = new GraphAuth(clientId, ini.Get("graph", "tenant", "consumers"));
        var token = await auth.TryGetTokenSilentAsync(CancellationToken.None);
        if (token is null) { sb.AppendLine("NOT signed in (tray menu → Sign in to Graph)"); return Write(outPath, sb); }

        var snap = await new GraphClient(_ => Task.FromResult(token)).FetchAsync(CancellationToken.None);
        var index = new SearchIndexReader(); index.Reload();
        var grace = TimeSpan.FromMinutes(ini.GetInt("general", "grace_minutes", 10));
        var now = DateTimeOffset.UtcNow;

        sb.AppendLine($"Graph: {snap.Notebooks.Count} notebooks, {snap.Sections.Count} sections.  Local index: {index.Notebooks().Count()} notebooks, {index.Sections().Count()} sections.  grace={grace.TotalMinutes} min");
        sb.AppendLine();
        sb.AppendLine("GRAPH NOTEBOOKS");
        foreach (var nb in snap.Notebooks) sb.AppendLine($"  {nb.DisplayName,-30} lastModified={nb.LastModified:u}  id={nb.Id}");
        sb.AppendLine();
        sb.AppendLine("LOCAL INDEX NOTEBOOKS");
        foreach (var nb in index.Notebooks()) sb.AppendLine($"  {nb.Title,-30} newestInTree={index.NewestIn(nb):u}  gosid={nb.Gosid}");
        sb.AppendLine();
        sb.AppendLine("SECTIONS  (local newest vs server; + = local ahead, - = server ahead)");
        sb.AppendLine($"  {"notebook / section",-44} {"local newest (UTC)",-22} {"server (UTC)",-22} {"delta",-12} verdict");
        foreach (var gs in snap.Sections.OrderBy(s => s.NotebookName).ThenBy(s => s.DisplayName))
        {
            var local = index.Sections().FirstOrDefault(s => string.Equals(s.Title, gs.DisplayName, StringComparison.OrdinalIgnoreCase)
                && (gs.NotebookName is null || string.Equals(index.NotebookNameOf(s), gs.NotebookName, StringComparison.OrdinalIgnoreCase)))
                ?? index.Sections().FirstOrDefault(s => string.Equals(s.Title, gs.DisplayName, StringComparison.OrdinalIgnoreCase));
            var localNewest = local is null ? null : index.NewestIn(local);
            string verdict;
            if (local is null) verdict = "no local match";
            else if (localNewest is null || gs.LastModified is null) verdict = "missing timestamp";
            else
            {
                var d = localNewest.Value - gs.LastModified.Value;
                verdict = d > grace ? "local ahead (index stamp; alerts only if it moves again while watching)"
                        : d < -(grace + TimeSpan.FromMinutes(5)) ? "server ahead (informational)" : "in sync";
            }
            var delta = localNewest is not null && gs.LastModified is not null ? $"{(localNewest.Value - gs.LastModified.Value).TotalMinutes:+0;-0} min" : "";
            sb.AppendLine($"  {gs.NotebookName + " / " + gs.DisplayName,-44} {localNewest?.ToString("u") ?? "-",-22} {gs.LastModified?.ToString("u") ?? "-",-22} {delta,-12} {verdict}");
            sb.AppendLine($"      graph id={gs.Id}");
            if (gs.ClientUrl is not null) sb.AppendLine($"      clientUrl={gs.ClientUrl}");
        }
        sb.AppendLine();
        var det = new OutcomeDetector(grace);
        var issues = det.Evaluate(index, snap, oneNoteRunning: true, internet: true, now);
        sb.AppendLine($"OUTCOME DETECTOR: baselined {det.BaselinedSections} sections on this first poll; would raise now: {issues.Count}");
        sb.AppendLine("  (an UploadStuck alert needs a local change observed on a later poll that the server does not pick up within grace;");
        sb.AppendLine("   the search index re-stamps LastModifiedTime when a notebook is (re)opened, so absolute skew above is NOT a failure)");
        foreach (var i in issues) { sb.AppendLine(); sb.AppendLine(i.ToReport()); }
        return Write(outPath, sb);
    }

    private static string Write(string path, StringBuilder sb) { File.WriteAllText(path, sb.ToString()); return path; }
}
