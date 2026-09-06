using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher;

/// <summary>Everything the window needs for one render.</summary>
public sealed record IssuesView(IReadOnlyList<SyncIssue> Issues, WatcherStatus? Status, Func<SyncIssue, bool> IsIgnored, GraphPoller? Graph);

/// <summary>
/// Live "issues &amp; status" window: STATUS, then every active issue as WHERE / WHAT / WHY / FIX.
/// Non-modal, self-refreshing every 3 s. Graph sign-in/out lives here so the action is right where
/// the "not signed in" line appears.
/// </summary>
public sealed class IssuesForm : Form
{
    private readonly TextBox _text = new()
    {
        Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Dock = DockStyle.Fill,
        Font = new Font("Consolas", 10f), WordWrap = true,
    };
    private readonly FlowLayoutPanel _bar = new() { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(6, 6, 6, 2) };
    private readonly Button _signIn = new() { Text = "Sign in to Graph…", AutoSize = true };
    private readonly Button _signOut = new() { Text = "Sign out of Graph", AutoSize = true };
    private readonly Button _logs = new() { Text = "Open logs folder", AutoSize = true };
    private readonly Label _footer = new() { Dock = DockStyle.Bottom, Height = 22, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(6, 0, 0, 0) };
    private readonly System.Windows.Forms.Timer _timer = new() { Interval = 3000 };
    private readonly Func<IssuesView> _snapshot;
    private string _last = "";

    public IssuesForm(Func<IssuesView> snapshot, Func<Task> signIn, Func<Task> signOut, Action openLogs)
    {
        _snapshot = snapshot;
        Text = "OneNote Sync Watcher — issues & status";
        Width = 940; Height = 620; StartPosition = FormStartPosition.CenterScreen;

        _signIn.Click += async (_, _) => { _signIn.Enabled = false; try { await signIn(); } finally { _signIn.Enabled = true; Render(); } };
        _signOut.Click += async (_, _) => { await signOut(); Render(); };
        _logs.Click += (_, _) => openLogs();
        _bar.Controls.AddRange([_signIn, _signOut, _logs]);

        Controls.Add(_text); Controls.Add(_bar); Controls.Add(_footer);
        _timer.Tick += (_, _) => Render();
        Shown += (_, _) => { Render(); _timer.Start(); };
        FormClosed += (_, _) => _timer.Stop();
    }

    public void Render()
    {
        var v = _snapshot();
        var graphOn = v.Graph is not null;
        _signIn.Visible = graphOn && !v.Graph!.SignedIn;
        _signOut.Visible = graphOn && v.Graph!.SignedIn;

        var text = Build(v);
        _footer.Text = $"auto-refreshes every 3 s — last update {DateTime.Now:HH:mm:ss}";
        if (text == _last) return;
        _last = text;
        var first = SendMessage(_text.Handle, EM_GETFIRSTVISIBLELINE, 0, 0);
        _text.Text = text;
        _text.Select(0, 0);
        SendMessage(_text.Handle, EM_LINESCROLL, 0, first);
    }

    private static string Build(IssuesView v)
    {
        var sb = new System.Text.StringBuilder();
        var nl = Environment.NewLine;
        var s = v.Status;

        // Two different shapes, because the data has two different shapes. STATUS is key/value, so it
        // gets a narrow fixed grid whose width depends only on its own five labels. NOTEBOOKS is
        // tabular, so it gets a real table with a header. Merging them let one long notebook name
        // shove the whole status column right, which is what made the window look ragged.
        const string pad = "  ";
        var width = new[] { "Graph", "Collector", "OneNote", "Last telemetry", "Last sync" }.Max(l => l.Length) + 2;
        string Row(string label, string value) => $"{pad}{(label + ":").PadRight(width)}{value}{nl}";

        sb.Append(pad).Append(v.Issues.Count == 0
            ? "SUCCESS — nothing to look at."
            : $"ERROR — {v.Issues.Count} problem(s) need attention").Append(nl).Append(nl);

        sb.Append("STATUS").Append(nl);
        if (v.Graph is not null)
        {
            var graph = v.Graph.SignedIn ? "signed in" : "NOT signed in — use the button above";
            if (v.Graph.LastError is not null) graph += $" (last error: {v.Graph.LastError})";
            else if (v.Graph.LastFetchUtc is { } g) graph += $" (last check {g.ToLocalTime():HH:mm})";
            sb.Append(Row("Graph", graph));
        }

        if (s is null)
        {
            sb.Append(Row("Collector", "NOT running — no status file"));
        }
        else
        {
            var age = DateTimeOffset.UtcNow - s.UpdatedUtc;
            var alive = s.CollectorRunning && age < TimeSpan.FromMinutes(2);
            sb.Append(Row("Collector", $"{(alive ? "running" : "NOT running")} ({age.TotalSeconds:F0}s ago)"));
            sb.Append(Row("OneNote", s.OneNoteRunning ? "running" : "not running"));
            sb.Append(Row("Last telemetry", Time(s.LastTelemetryUtc, "HH:mm:ss")));
            sb.Append(Row("Last sync", Time(s.LastSyncEventUtc, "HH:mm:ss")));

            if (s.Notebooks.Count > 0)
            {
                // Column widths come from the content, but the name column is capped so one very long
                // notebook name cannot stretch the table past the window.
                var rows = s.Notebooks
                    .Select(nb => (Name: Ellipsis(nb.Label, 36), State: nb.LastSuccess switch
                    {
                        true => "OK",
                        false => $"FAILED {nb.LastCode} {nb.LastError}".Trim(),
                        _ => "unknown",
                    }, Seen: Time(nb.LastSyncUtc, "HH:mm:ss")))
                    .Select(r => (r.Name, State: Ellipsis(r.State, 40), r.Seen))
                    .ToList();
                var nameCol = Math.Max(8, rows.Max(r => r.Name.Length)) + 2;
                var stateCol = Math.Max(5, rows.Max(r => r.State.Length)) + 2;

                sb.Append(nl).Append("NOTEBOOKS").Append(nl);
                sb.Append(pad).Append("NOTEBOOK".PadRight(nameCol)).Append("STATE".PadRight(stateCol))
                  .Append("LAST ACTIVITY").Append(nl);
                foreach (var r in rows)
                    sb.Append(pad).Append(r.Name.PadRight(nameCol)).Append(r.State.PadRight(stateCol))
                      .Append(r.Seen).Append(nl);
            }
        }

        sb.Append(nl)
          .Append(pad).Append("Last telemetry is the pipeline-alive signal. A quiet Last sync is").Append(nl)
          .Append(pad).Append("normal — an idle OneNote can go ~2 h between syncs.").Append(nl);

        if (v.Issues.Count > 0)
        {
            sb.Append(nl).Append("PROBLEMS").Append(nl);
            foreach (var i in v.Issues.OrderByDescending(i => i.LastSeen))
            {
                sb.Append(nl).Append($"  ── from {i.Detector}{(v.IsIgnored(i) ? " (ignored by config)" : "")}").Append(nl);
                foreach (var line in i.ToReport().Split(Environment.NewLine))
                    sb.Append("  ").Append(line).Append(nl);
            }
        }
        return sb.ToString();
    }

    /// <summary>Truncate for a fixed table column, so one long name cannot stretch the table.</summary>
    private static string Ellipsis(string s, int max) => s.Length <= max ? s : s[..(max - 1)] + "…";

    private static string Time(DateTimeOffset? t, string format) =>
        t is { } v ? v.ToLocalTime().ToString(format) : "none yet";

    private const int EM_GETFIRSTVISIBLELINE = 0x00CE;
    private const int EM_LINESCROLL = 0x00B6;
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, int lParam);

    protected override void Dispose(bool disposing) { if (disposing) _timer.Dispose(); base.Dispose(disposing); }
}
