using System.Text;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Collector;

/// <summary>
/// <c>--audit</c>: runs every OneNote SendEvent found in the Office diagnostic logs through the real
/// classifier and reports how each event name was classified. This is the evidence that the
/// fail-closed contract holds on real data — no event is silently dropped, and nothing is called a
/// success without positive proof. See docs/fail-closed.md.
/// </summary>
public static class ClassificationAudit
{
    public static string Run(string diagDir)
    {
        var counts = new Dictionary<string, Dictionary<SyncOutcome, int>>(StringComparer.Ordinal);
        var ignored = new Dictionary<string, int>(StringComparer.Ordinal);
        int lines = 0, oneNoteLines = 0, malformed = 0;

        foreach (var file in Directory.Exists(diagDir) ? Directory.EnumerateFiles(diagDir, "Primary*.log") : [])
        {
            string text;
            try
            {
                using var fs = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var sr = new StreamReader(fs);
                text = sr.ReadToEnd().Replace("\0", "");
            }
            catch (IOException) { continue; }   // live session, still locked

            foreach (var line in text.Split('\n'))
            {
                var i = line.IndexOf("SendEvent {", StringComparison.Ordinal);
                if (i < 0) continue;
                lines++;
                var msg = line[i..].TrimEnd('\r', '\t');
                if (!msg.Contains("\"Office.OneNote.", StringComparison.Ordinal)) continue;
                oneNoteLines++;

                var ev = SyncEventJson.TryParse(msg, out var bad);
                if (bad) malformed++;
                if (ev is null)
                {
                    var name = NameOf(msg) ?? "(unknown)";
                    ignored[name] = ignored.GetValueOrDefault(name) + 1;
                    continue;
                }
                if (!counts.TryGetValue(ev.EventName, out var byOutcome))
                    counts[ev.EventName] = byOutcome = new Dictionary<SyncOutcome, int>();
                byOutcome[ev.Outcome] = byOutcome.GetValueOrDefault(ev.Outcome) + 1;
            }
        }

        var sb = new StringBuilder();
        sb.AppendLine($"Classification audit — {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss}");
        sb.AppendLine($"source: {diagDir}");
        sb.AppendLine($"SendEvent rows: {lines};  Office.OneNote.* rows: {oneNoteLines};  malformed OneNote payloads: {malformed}");
        sb.AppendLine();
        sb.AppendLine("IN SCOPE — every event below was classified (a 'Success' requires positive proof):");
        sb.AppendLine($"  {"event",-58} {"count",6}  outcomes");
        foreach (var kv in counts.OrderBy(k => k.Key, StringComparer.Ordinal))
        {
            var total = kv.Value.Values.Sum();
            var detail = string.Join(", ", kv.Value.OrderByDescending(o => o.Value).Select(o => $"{o.Key}={o.Value}"));
            sb.AppendLine($"  {kv.Key,-58} {total,6}  {detail}");
        }
        sb.AppendLine();
        sb.AppendLine("DELIBERATELY OUT OF SCOPE — no failure signal in name or fields (not sync outcomes):");
        foreach (var kv in ignored.OrderByDescending(k => k.Value))
            sb.AppendLine($"  {kv.Key,-58} {kv.Value,6}");
        sb.AppendLine();
        var problems = counts.Sum(k => k.Value.Where(o => o.Key is SyncOutcome.Failure or SyncOutcome.Transient
            or SyncOutcome.SuspectedFailure or SyncOutcome.Unknown).Sum(o => o.Value));
        sb.AppendLine($"events that would have surfaced as a problem: {problems}");
        return sb.ToString();
    }

    private static string? NameOf(string msg)
    {
        var m = System.Text.RegularExpressions.Regex.Match(msg, "\"EventName\"\\s*:\\s*\"([^\"]+)\"");
        return m.Success ? m.Groups[1].Value : null;
    }
}
