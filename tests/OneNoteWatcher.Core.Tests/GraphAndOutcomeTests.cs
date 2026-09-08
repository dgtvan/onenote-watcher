using Microsoft.Data.Sqlite;
using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

public class SectionNameMapTests
{
    private static GraphSnapshot Snap() => new(DateTimeOffset.UtcNow,
        [new GraphNotebook("nb1", "Work", null)],
        // real id shape observed on 2026-09-06: "0-<drive>!s<32 hex>" — embeds the OneDrive resource id
        [new GraphSection("0-36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2", "ASW", null, "nb1", "Work",
            "onenote:https://d.docs.live.net/36B934175DC7E3A4/Work/ASW.one#section-id={F871B437-99B3-4241-A3F4-FFCBDE99566E}&end",
            "https://onedrive.live.com/redir?resid=x&page=Edit&wd=target%28ASW.one%7C..%29")]);

    [Fact]
    public void Resolves_by_resource_id_token_and_by_section_guid()
    {
        var map = SectionNameMap.FromSnapshot(Snap());
        Assert.Equal("Work / ASW", map.Lookup("36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2", null));
        Assert.Equal("Work / ASW", map.Lookup(null, "{F871B437-99B3-4241-A3F4-FFCBDE99566E}{1}"));
        Assert.Null(map.Lookup("36B934175DC7E3A4!sdeadbeefdeadbeefdeadbeefdeadbeef", null));
    }

    [Fact]
    public void A_section_with_a_numeric_item_id_resolves_from_a_sync_events_spelling()
    {
        // "0-<drive>!<number>" — the shape the old lookup could not match, so the history showed the raw
        // id ("Note / …DC7E3A4!1242") instead of the section name
        var snap = new GraphSnapshot(DateTimeOffset.UtcNow,
            [new GraphNotebook("nb1", "Note", null)],
            [new GraphSection("0-36B934175DC7E3A4!1242", "Family", null, "nb1", "Note", null, null)]);

        Assert.Equal("Note / Family", SectionNameMap.FromSnapshot(snap).Lookup("36B934175DC7E3A4!1242", null));
    }

    [Fact]
    public void The_shared_drive_id_is_never_a_key_so_sections_cannot_be_confused()
    {
        // every section in a drive repeats the same drive id; keying on it made them overwrite one
        // another, and a lookup then returned a confidently wrong name
        var snap = new GraphSnapshot(DateTimeOffset.UtcNow,
            [new GraphNotebook("nb1", "Note", null)],
            [new GraphSection("0-36B934175DC7E3A4!1242", "Family", null, "nb1", "Note", null, null),
             new GraphSection("0-36B934175DC7E3A4!943", "Random", null, "nb1", "Note", null, null)]);
        var map = SectionNameMap.FromSnapshot(snap);

        Assert.Null(map.Lookup("36B934175DC7E3A4", null));                    // ambiguous → no answer
        Assert.Equal("Note / Family", map.Lookup("36B934175DC7E3A4!1242", null));
        Assert.Equal("Note / Random", map.Lookup("36B934175DC7E3A4!943", null));
    }

    [Fact]
    public void Round_trips_through_file()
    {
        var path = Path.Combine(Path.GetTempPath(), "onwatch_secmap_" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            SectionNameMap.FromSnapshot(Snap()).Save(path);
            Assert.Equal("Work / ASW", SectionNameMap.Load(path).Lookup("x!s8d499fcb43aa42bd9a75d6d555478ac2", null));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void Resolver_splits_notebook_and_section_from_map_value()
    {
        var map = SectionNameMap.FromSnapshot(Snap());
        var r = new NameResolver(new SearchIndexReader(Path.GetTempPath() + "\\nope"), new MruReader(Path.GetTempPath() + "\\nope"))
        { SectionNameByResourceId = map.Lookup };
        var ev = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.SectionResourceId_ResourceId":"36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2","Data.Success":false,"Data.Error_Code":3758096477}""")!;
        Assert.Equal(("Work", "ASW"), (r.Resolve(ev).Notebook, r.Resolve(ev).Section));
    }

    [Fact]
    public void Resolver_prefers_mru_display_name_and_learns_ids_both_ways()
    {
        var dir = Path.Combine(Path.GetTempPath(), "onwatch_learn_" + Guid.NewGuid().ToString("N"));
        var idxDir = Path.Combine(dir, "idx"); Directory.CreateDirectory(idxDir);
        var mruDir = Path.Combine(dir, "mru", "id_LiveId", "OneNote"); Directory.CreateDirectory(mruDir);
        using (var con = new SqliteConnection($"Data Source={Path.Combine(idxDir, "nb.db")}"))
        {
            con.Open();
            using var cmd = con.CreateCommand();
            cmd.CommandText = "CREATE TABLE Entities (Type INTEGER, GOID TEXT, GOSID TEXT, ParentGOID TEXT, Title TEXT, LastModifiedTime INTEGER);" +
                              "INSERT INTO Entities VALUES (4,'{N}{1}','{F282F88E-ED0A-4BF0-BD60-F9D9ECAC9A33}{1}',NULL,'Van',0);";
            cmd.ExecuteNonQuery();
        }
        File.WriteAllText(Path.Combine(mruDir, "Documents_en-US"), """[{"FileName":"Note","ResourceId":"36b934175dc7e3a4!626"}]""");
        try
        {
            var r = new NameResolver(new SearchIndexReader(idxDir), new MruReader(Path.Combine(dir, "mru")));
            r.Refresh(TimeSpan.Zero);
            // gosid-only event before any learning → index title is the only source
            var gosidOnly = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Gosid":"{F282F88E-ED0A-4BF0-BD60-F9D9ECAC9A33}{1}","Data.Success":true}""")!;
            Assert.Equal("Van", r.Resolve(gosidOnly).Notebook);
            // learn rid<->gosid from a full notebook event → original name "Note" + display name "Van"
            r.Learn(SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Gosid":"{F282F88E-ED0A-4BF0-BD60-F9D9ECAC9A33}{1}","Data.NotebookId_ResourceId":"36B934175DC7E3A4!626","Data.Success":true}""")!);
            var n = r.Resolve(gosidOnly);
            Assert.Equal(("Note", "Van"), (n.Notebook, n.NotebookDisplayName));
            Assert.Equal("Note (Display name: Van)", n.NotebookLabel);
            var ridOnly = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.NotebookId_ResourceId":"36B934175DC7E3A4!626","Data.Success":true}""")!;
            Assert.Equal("Note (Display name: Van)", r.Resolve(ridOnly).NotebookLabel);
        }
        finally { try { Directory.Delete(dir, true); } catch { } }
    }
}

public class OutcomeDetectorTests
{
    private static readonly DateTimeOffset T0 = new(2026, 9, 6, 10, 0, 0, TimeSpan.Zero);
    private static readonly TimeSpan Grace = TimeSpan.FromMinutes(10);

    private static IReadOnlyList<LocalSection> Local(DateTimeOffset asw, DateTimeOffset cred) =>
        [new LocalSection("Work", "ASW", asw), new LocalSection("Work", "Credentials", cred)];

    private static GraphSnapshot Server(DateTimeOffset asw, DateTimeOffset cred) => new(T0,
        [new GraphNotebook("nb", "Work", null)],
        [new GraphSection("gs1", "ASW", asw, "nb", "Work", null, null),
         new GraphSection("gs2", "Credentials", cred, "nb", "Work", null, null)]);

    [Fact]
    public void Preexisting_skew_at_startup_never_alerts()
    {
        // the real false-positive case: index stamps months ahead of Graph, nothing changing
        var det = new OutcomeDetector(Grace);
        var local = Local(T0.AddDays(-2), T0.AddDays(-30));
        var server = Server(T0.AddDays(-12), T0.AddDays(-120));
        Assert.Empty(det.Evaluate(local, server, true, true, T0));
        Assert.Empty(det.Evaluate(local, server, true, true, T0.AddMinutes(15)));
        Assert.Empty(det.Evaluate(local, server, true, true, T0.AddHours(3)));
        Assert.Equal(2, det.BaselinedSections);
    }

    [Fact]
    public void Change_while_watching_that_server_never_picks_up_alerts_after_grace()
    {
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);

        // user edits ASW at T0+1m; server does not move
        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
        Assert.Empty(det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0.AddMinutes(2)));   // within grace
        var issues = det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0.AddMinutes(13));   // grace passed
        var i = Assert.Single(issues);
        Assert.Equal(IssueKind.UploadStuck, i.Kind);
                Assert.Equal("Work / ASW", i.Location);
        Assert.Equal(T0.AddMinutes(1), i.EvidenceUtc);
        Assert.Contains("Shift+F9", i.Recommendation);
    }

    [Fact]
    public void Change_that_server_picks_up_clears_and_rebaselines()
    {
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
        det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0.AddMinutes(2));
        // server catches up (30 s after local, as observed in real data)
        Assert.Empty(det.Evaluate(edited, Server(T0.AddMinutes(1.5), T0.AddDays(-120)), true, true, T0.AddMinutes(13)));
        Assert.Empty(det.Evaluate(edited, Server(T0.AddMinutes(1.5), T0.AddDays(-120)), true, true, T0.AddMinutes(30)));
    }

    [Fact]
    public void Server_slightly_before_local_still_counts_as_synced()
    {
        // real data: "Links" local 15:03:14 vs server 15:03:01 while perfectly synced
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
        Assert.Empty(det.Evaluate(edited, Server(T0.AddMinutes(1).AddSeconds(-13), T0.AddDays(-120)), true, true, T0.AddMinutes(20)));
    }

    [Fact]
    public void Offline_downgrades_to_warning_with_network_category()
    {
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
        det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, false, T0.AddMinutes(2));
        var i = Assert.Single(det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, false, T0.AddMinutes(13)));
                Assert.Equal(FailureCategory.Network, i.Category);
    }

    [Fact]
    public void Server_only_movement_is_informational_not_an_alert()
    {
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
        var i = Assert.Single(det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddMinutes(-1)), true, true, T0.AddMinutes(5)));
        Assert.Equal(IssueKind.DownloadStuck, i.Kind);
            }

    [Fact]
    public void Edit_then_close_OneNote_leaves_stranded_changes_and_that_is_an_ALERT()
    {
        // OneNote emits no final sync on shutdown (verified on real sessions), so an edit made just
        // before closing can leave NO telemetry. The cloud-side check is the only thing that catches it,
        // and with OneNote closed nothing will retry — so it must be an alert, not a quiet warning.
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);

        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));           // user edits at T0+1
        det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), oneNoteRunning: true, internet: true, T0.AddMinutes(2));

        // OneNote is now CLOSED and the server never picked the change up
        var issues = det.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)),
            oneNoteRunning: false, internet: true, T0.AddMinutes(13));
        var i = Assert.Single(issues);
        Assert.Equal(IssueKind.UploadStuck, i.Kind);
                Assert.Equal("Work / ASW", i.Location);
        Assert.Contains("not uploaded before OneNote was closed", i.Summary);
        Assert.Contains("Start OneNote", i.Recommendation);
    }

    [Fact]
    public void Closing_OneNote_with_everything_synced_raises_nothing()
    {
        var det = new OutcomeDetector(Grace);
        det.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
        var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
        det.Evaluate(edited, Server(T0.AddMinutes(1.5), T0.AddDays(-120)), true, true, T0.AddMinutes(2));   // server caught up
        Assert.Empty(det.Evaluate(edited, Server(T0.AddMinutes(1.5), T0.AddDays(-120)), false, true, T0.AddMinutes(30)));
    }

    [Fact]
    public void Baseline_survives_a_restart_so_a_change_stranded_before_reboot_is_still_flagged()
    {
        // edit → close OneNote → shut down the PC. On next login the detector must NOT re-baseline the
        // stranded change as normal, or it would be silently lost.
        var path = Path.Combine(Path.GetTempPath(), "onwatch_state_" + Guid.NewGuid().ToString("N") + ".json");
        try
        {
            var before = new OutcomeDetector(Grace);
            before.Evaluate(Local(T0.AddDays(-2), T0.AddDays(-30)), Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0);
            var edited = Local(T0.AddMinutes(1), T0.AddDays(-30));
            before.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), true, true, T0.AddMinutes(2));
            before.SaveState(path);           // ... and the PC shuts down

            var afterReboot = new OutcomeDetector(Grace);
            Assert.Equal(2, afterReboot.LoadState(path));
            // next day: OneNote closed, the change still never reached OneDrive
            var i = Assert.Single(afterReboot.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)),
                oneNoteRunning: false, internet: true, T0.AddDays(1)));
            Assert.Equal(IssueKind.UploadStuck, i.Kind);
                        // a fresh detector with no saved state would have silently accepted it
            var naive = new OutcomeDetector(Grace);
            Assert.Empty(naive.Evaluate(edited, Server(T0.AddDays(-12), T0.AddDays(-120)), false, true, T0.AddDays(1)));
        }
        finally { File.Delete(path); }
    }
}
