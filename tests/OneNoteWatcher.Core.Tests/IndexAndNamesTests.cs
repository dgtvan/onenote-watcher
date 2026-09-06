using Microsoft.Data.Sqlite;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

/// <summary>Builds a tiny FullTextSearchIndex-shaped SQLite db + MRU file and checks name resolution.</summary>
public class IndexAndNamesTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "onwatch_idx_" + Guid.NewGuid().ToString("N"));
    private readonly string _idxDir, _mruDir;

    private const string NbGosid = "{42AC7CD5-3122-4501-B9DF-254D9AC6D8F6}{1}";
    private const string NbGoid = "{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{10}";
    private const string SecGoid = "{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{32}";
    private const string PageGoid = "{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{108}";
    private const string NbRid = "36b934175dc7e3a4!s6f96513252744f6e8de2254341eb06e0";

    public IndexAndNamesTests()
    {
        _idxDir = Path.Combine(_dir, "FullTextSearchIndex"); Directory.CreateDirectory(_idxDir);
        _mruDir = Path.Combine(_dir, "Mru", "id_LiveId", "OneNote"); Directory.CreateDirectory(_mruDir);

        var db = Path.Combine(_idxDir, "{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{10}.db");
        using var con = new SqliteConnection($"Data Source={db}");
        con.Open();
        using var cmd = con.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE Entities (rowid INTEGER PRIMARY KEY, Type INTEGER, GOID TEXT, GUID TEXT, GOSID TEXT,
              ParentGOID TEXT, GrandparentGOIDs TEXT, ContentRID TEXT, RootRevGenCount INTEGER,
              LastModifiedTime INTEGER, RecentTime INTEGER, PinTime INTEGER, Color INTEGER, Title TEXT, EnterpriseIdentity TEXT);
            INSERT INTO Entities(Type,GOID,GUID,GOSID,ParentGOID,LastModifiedTime,Title) VALUES
              (4,'{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{10}','g',  '{42AC7CD5-3122-4501-B9DF-254D9AC6D8F6}{1}', NULL, 133727000000000000,'Work'),
              (2,'{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{32}','g2', '{53670BD0-CBA0-0830-0B3B-CC82BA7229B2}{1}','{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{10}',133727000000000000,'ASW'),
              (1,'{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{108}','g3','{3A72756B-DD0C-7B87-B539-0D64DA7A97D1}{1}','{7B3AFF82-9493-454E-9AC1-3331ABF2826A}{32}',133727100000000000,'Eyes check');
            """;
        cmd.ExecuteNonQuery();

        File.WriteAllText(Path.Combine(_mruDir, "Documents_en-US"),
            $$"""[{"FileName":"Work","ResourceId":"{{NbRid}}","DocumentUrl":"https://d.docs.live.net/36B934175DC7E3A4/Work"}]""");
    }

    [Fact]
    public void Index_resolves_notebook_section_page_and_newest_time()
    {
        var idx = new SearchIndexReader(_idxDir); idx.Reload();
        var nb = idx.ByGosid(NbGosid + "{B0}"); // COM-style suffix tolerated
        Assert.Equal("Work", nb!.Title);
        var page = idx.ByGoid(PageGoid)!;
        Assert.Equal("Eyes check", page.Title);
        Assert.Equal("ASW", idx.SectionNameOf(page));
        Assert.Equal("Work", idx.NotebookNameOf(page));
        var newest = idx.NewestIn(idx.ByGoid(NbGoid)!);
        Assert.Equal(DateTimeOffset.FromFileTime(133727100000000000).ToUniversalTime(), newest);
    }

    [Fact]
    public void Resolver_names_from_index_mru_and_page_goid()
    {
        var resolver = new NameResolver(new SearchIndexReader(_idxDir), new MruReader(Path.Combine(_dir, "Mru")));
        resolver.Refresh(TimeSpan.Zero);

        var nbEvent = SyncEventJson.TryParse($$"""SendEvent {"EventName":"Office.OneNote.Storage.NotebookSyncResult","Data.Gosid":"{{NbGosid}}","Data.Success":true}""")!;
        Assert.Equal("Work", resolver.Resolve(nbEvent).Notebook);

        var ridOnly = SyncEventJson.TryParse($$"""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.NotebookId_ResourceId":"{{NbRid}}","Data.Success":true}""")!;
        Assert.Equal("Work", resolver.Resolve(ridOnly).Notebook);

        var pg = SyncEventJson.TryParse($$"""SendEvent {"EventName":"Office.OneNote.Storage.PageSyncSession","Data.ActivePageGOID":"{{PageGoid}}"}""")!;
        var r = resolver.Resolve(pg);
        Assert.Equal(("Work", "ASW", "Eyes check"), (r.Notebook, r.Section, r.Page));

        resolver.SectionNameByResourceId = (rid, _) => rid is not null && rid.EndsWith("abc") ? "Graph Section" : null;
        var sec = SyncEventJson.TryParse("""SendEvent {"EventName":"Office.OneNote.Storage.SectionSyncResult","Data.SectionResourceId_ResourceId":"x!sabc","Data.Success":true}""")!;
        Assert.Equal("Graph Section", resolver.Resolve(sec).Section);
    }

    public void Dispose() { try { Directory.Delete(_dir, true); } catch { } }
}
