using OneNoteWatcher.Core;
using OneNoteWatcher.Core.Diagnosis;
using OneNoteWatcher.Core.History;
using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Parsing;

namespace OneNoteWatcher.Core.Tests;

public class SyncEventMapperTests
{
    private static SyncEvent Failure => SyncEventJson.TryParse(
        "SendEvent {\"EventName\":\"Office.OneNote.Storage.SectionSyncResult\"," +
        "\"Time\":\"2026-09-04T15:54:48.427Z\",\"Data.Success\":false," +
        "\"Data.Error_Code\":3758096477,\"Data.Error_Description\":\"ErrFilePendingRename\"," +
        "\"Data.IsErrorTransient\":false,\"Data.SectionResourceId_ResourceId\":\"36B934175DC7E3A4!se6abe06d75b4\"}")!;

    [Fact]
    public void Failure_maps_to_issue_with_where_what_why_fix()
    {
        var issue = SyncEventMapper.ToIssue(Failure, "etw", new ResolvedNames("Work", "ASW", null));
        Assert.NotNull(issue);
        Assert.Equal(IssueKind.ErrorCodeReported, issue!.Kind);
        Assert.Equal("0xE000005D", issue.Code);
        Assert.Equal("Work / ASW", issue.Location);
        Assert.Equal(FailureCategory.FileState, issue.Category);
        Assert.Contains("pending a rename", issue.Summary);
        Assert.Contains("close and reopen", issue.Recommendation);
        Assert.Contains("section-rid=36B934175DC7E3A4!se6abe06d75b4", issue.TechnicalDetail);
        var report = issue.ToReport();
        Assert.Contains("WHERE : Work / ASW", report);
        Assert.Contains("FIX   :", report);
    }

    [Fact]
    public void Transient_known_code_is_warn_not_alert()
    {
        var e = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.SectionSyncResult\",\"Data.Success\":false,\"Data.Error_Code\":3825336384}")!; // 0xE4020040 store busy
        var issue = SyncEventMapper.ToIssue(e, "etw");
        Assert.NotNull(issue);
                Assert.Equal(FailureCategory.Transient, issue.Category);
    }

    [Fact]
    public void SyncScore_is_diagnostic_only_never_an_issue()
    {
        // the real false alarm: OneNote's replication scan skipping a locked password-protected section
        var e = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.SyncScore\",\"Time\":\"2026-09-06T04:52:00Z\"," +
            "\"Data.Error_Code\":3758097184,\"Data.Error_Description\":\"ErrCrypto_BadPassphrase\",\"Data.Error_Type\":\"Win32Error\",\"Data.Source\":\"Storage.Replication.Fishbowl\"}")!;
        Assert.Null(SyncEventMapper.ToIssue(e, "etw"));
        var line = SyncEventMapper.ToHistoryLine(e);
        Assert.Contains("(replication scan)  SCORE  0xE0000320 ErrCrypto_BadPassphrase", line);
        Assert.Contains("background replication scan", line);   // says WHY it is not an alert
        Assert.DoesNotContain("FAILED", line);
    }

    [Fact]
    public void Locked_section_failure_is_a_warning_with_display_name_in_location()
    {
        var e = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.SectionSyncResult\",\"Data.Success\":false," +
            "\"Data.Error_Code\":3758097184,\"Data.Error_Description\":\"ErrCrypto_BadPassphrase\",\"Data.SectionResourceId_ResourceId\":\"x!s1\"}")!;
        var issue = SyncEventMapper.ToIssue(e, "etw", new ResolvedNames("Note", "Secrets", null, "Van"))!;
                Assert.Equal("Note (Display name: Van) / Secrets", issue.Location);
        Assert.Contains("unlock", issue.Recommendation, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Success_maps_to_no_issue()
    {
        var ok = SyncEventJson.TryParse(
            "SendEvent {\"EventName\":\"Office.OneNote.Storage.NotebookSyncResult\",\"Data.Success\":true}")!;
        Assert.Null(SyncEventMapper.ToIssue(ok, "etw"));
    }

    [Fact]
    public void History_line_has_scope_result_code_and_reason()
    {
        var line = SyncEventMapper.ToHistoryLine(Failure, new ResolvedNames("Work", "ASW", null));
        Assert.Contains("Work / ASW", line);
        Assert.Contains("FAILED", line);
        Assert.Contains("0xE000005D ErrFilePendingRename", line);
        Assert.Contains("[FileState]", line);
    }

    [Fact]
    public void History_line_falls_back_to_short_ids_when_names_unknown()
    {
        var line = SyncEventMapper.ToHistoryLine(Failure);
        Assert.Contains("FAILED", line);
        Assert.Contains("…", line); // shortened resource id
    }

    [Fact]
    public void Connectivity_and_page_session_lines()
    {
        var off = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.ConnectivityChanged\",\"Data.InternetConnectivityNowAvailable\":false}")!;
        Assert.Equal(SyncEventKind.ConnectivityChanged, off.Kind);
        Assert.False(off.InternetAvailable);
        Assert.Contains("OFFLINE", SyncEventMapper.ToHistoryLine(off));

        var pg = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.PageSyncSession\",\"Data.ActivePageGOID\":\"{BE7F798D-B55F-4C2D-A3A6-817E121F03E2}{108}\",\"Data.ErrorState_TimeInActiveSyncErrorState\":1200}")!;
        Assert.Equal("{BE7F798D-B55F-4C2D-A3A6-817E121F03E2}{108}", pg.PageGoid);
        Assert.Equal(1200, pg.TimeInSyncErrorStateMs);
        Assert.Equal(SyncOutcome.SuspectedFailure, pg.Outcome); // time in an error state is a signal, not benign
        Assert.Contains("PAGE-SESSION  SUSPECT  errorStateMs=1200", SyncEventMapper.ToHistoryLine(pg, new ResolvedNames("Note", "Quick Notes", "Eyes check")));

        var healthy = SyncEventJson.TryParse("SendEvent {\"EventName\":\"Office.OneNote.Storage.PageSyncSession\",\"Data.ErrorState_TimeInActiveSyncErrorState\":0}")!;
        Assert.Equal(SyncOutcome.Diagnostic, healthy.Outcome);
        Assert.Contains("PAGE-SESSION  OK", SyncEventMapper.ToHistoryLine(healthy));
    }
}

public class ErrorCatalogTests
{
    [Theory]
    [InlineData(0xE000005Du, FailureCategory.FileState)]
    [InlineData(0xE0000320u, FailureCategory.Password)]
    [InlineData(0xE0000796u, FailureCategory.Storage)]
    [InlineData(0xE4010641u, FailureCategory.Network)]
    [InlineData(0xE40105F9u, FailureCategory.ClientVersion)]
    [InlineData(0xE000012Eu, FailureCategory.Corruption)]
    public void Known_codes_have_category_and_fix(uint code, FailureCategory cat)
    {
        var ex = ErrorCatalog.Explain(code, null);
        Assert.Equal(cat, ex.Category);
        Assert.False(string.IsNullOrWhiteSpace(ex.Recommendation));
    }

    [Theory]
    [InlineData("ErrCrypto_BadPassphrase", FailureCategory.Password)]
    [InlineData("ErrHttpTimeout", FailureCategory.Network)]
    [InlineData("ErrAccessDenied", FailureCategory.Permission)]
    [InlineData("jerrcFileNodeFileCorrupt_FileNodeListChunkMalformedRanOffEnd", FailureCategory.Corruption)]
    public void Unknown_code_is_categorised_by_description(string desc, FailureCategory cat)
    {
        Assert.Equal(cat, ErrorCatalog.Explain(0xE0009999, desc).Category);
    }

    [Fact]
    public void Totally_unknown_still_gets_actionable_text_and_reference()
    {
        var ex = ErrorCatalog.Explain(0xE4123456, null);
        Assert.Equal(FailureCategory.Service, ex.Category);
        Assert.Contains("0xE4123456", ex.Summary);
        Assert.Contains("Shift+F9", ex.Recommendation);
        Assert.NotNull(ex.Reference);
    }

    [Fact]
    public void Upload_stuck_explanations_depend_on_context()
    {
        Assert.Equal(FailureCategory.Network, ErrorCatalog.UploadStuck(true, false).Category);
        Assert.Contains("Start OneNote", ErrorCatalog.UploadStuck(false, true).Recommendation);
        Assert.Contains("Shift+F9", ErrorCatalog.UploadStuck(true, true).Recommendation);
    }
}

