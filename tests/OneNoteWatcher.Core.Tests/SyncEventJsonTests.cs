using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

public class SyncEventJsonTests
{
    private static string FixtureDir =>
        Path.Combine(AppContext.BaseDirectory, "fixtures");

    private static IEnumerable<string> Lines(string file) =>
        File.ReadAllLines(Path.Combine(FixtureDir, file))
            .Where(l => !string.IsNullOrWhiteSpace(l));

    [Fact]
    public void Parses_all_live_etw_fixtures()
    {
        var events = Lines("etw_sendevents.jsonl").Select(SyncEventJson.TryParse).ToList();
        Assert.NotEmpty(events);
        Assert.All(events, e => Assert.NotNull(e));
        // captured window was healthy → all successes, no failures
        Assert.All(events, e => Assert.Equal(SyncOutcome.Success, e!.Outcome));
    }

    [Fact]
    public void Decodes_real_failure_from_diaglog_fixture()
    {
        var events = Lines("diaglog_sendevents.jsonl").Select(SyncEventJson.TryParse).ToList();
        Assert.All(events, e => Assert.NotNull(e));

        var pendingRename = events.Single(e =>
            e!.Kind == SyncEventKind.SectionSyncResult && e.Success == false);
        Assert.Equal(SyncOutcome.Failure, pendingRename!.Outcome);
        Assert.Equal(0xE000005Du, pendingRename.ErrorCode);
        Assert.Equal("0xE000005D", pendingRename.ErrorCodeHex);
        Assert.Equal("ErrFilePendingRename", pendingRename.ErrorDescription);
        Assert.Equal("Win32Error", pendingRename.ErrorType);

        var badPass = events.Single(e => e!.Kind == SyncEventKind.SyncScore);
        Assert.Equal(0xE0000320u, badPass!.ErrorCode);
        Assert.Equal("ErrCrypto_BadPassphrase", badPass.ErrorDescription);
    }

    [Fact]
    public void Recovered_section_after_failure_is_not_a_failure()
    {
        // second SectionSyncResult on the same section succeeded
        var recovered = Lines("diaglog_sendevents.jsonl")
            .Select(SyncEventJson.TryParse)
            .Where(e => e!.Kind == SyncEventKind.SectionSyncResult && e.Success == true)
            .ToList();
        Assert.NotEmpty(recovered);
        Assert.All(recovered, e => Assert.Equal(SyncOutcome.Success, e!.Outcome));
    }

    [Theory]
    [InlineData("SendEvent {\"EventName\": \"Office.OneNote.Storage.NotebookSyncResult\", \"Data.Success\": true}", true)]
    [InlineData("{\"EventName\": \"Office.OneNote.Storage.SyncScore\", \"Data.Error_Code\": 3758096477}", true)]
    [InlineData("SendEvent {\"EventName\": \"Office.OneNote.Navigation.Navigate\"}", false)]
    [InlineData("SendEvent {\"EventName\": \"Office.Licensing.Something\"}", false)]
    [InlineData("not json at all", false)]
    [InlineData("", false)]
    [InlineData("SendEvent {truncated", false)]
    public void TryParse_accepts_only_storage_sync_events(string message, bool expected)
    {
        Assert.Equal(expected, SyncEventJson.TryParse(message) is not null);
    }

    [Fact]
    public void LooksLikeSyncSendEvent_prefilter_matches_parse()
    {
        // pre-filter must never reject something TryParse would accept
        foreach (var line in Lines("etw_sendevents.jsonl"))
        {
            var wz = "SendEvent " + line; // emulate wzMessage
            Assert.True(SyncEventJson.LooksLikeSyncSendEvent(wz));
            Assert.NotNull(SyncEventJson.TryParse(wz));
        }
        Assert.False(SyncEventJson.LooksLikeSyncSendEvent("SendEvent {\"EventName\":\"Office.Text.Foo\"}"));
    }

    [Fact]
    public void Parses_utc_time()
    {
        var e = SyncEventJson.TryParse(
            "SendEvent {\"EventName\":\"Office.OneNote.Storage.NotebookSyncResult\",\"Time\":\"2026-09-06T01:26:02Z\",\"Data.Success\":true}");
        Assert.NotNull(e);
        Assert.Equal(new DateTimeOffset(2026, 9, 6, 1, 26, 2, TimeSpan.Zero), e!.Time);
    }
}
