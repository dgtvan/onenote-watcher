using OneNoteWatcher.Core.Index;
using OneNoteWatcher.Core.Model;

namespace OneNoteWatcher.Collector;

/// <summary>
/// What a sync event is ABOUT, and what a successful one proves.
///
/// An error is only cleared by a success that genuinely COVERS it — same content, and later in time.
/// The coverage rules are deliberately narrow, because a wrongly-cleared error is a lost sync failure:
///
///  • A notebook sync success covers notebook-level failures ONLY. It does NOT cover its sections:
///    OneNote's own `NotebookSyncResult` carries separate `IsSectionErrorSuppressed` /
///    `IsSectionErrorUnexpected` fields, i.e. a notebook sync can report success while sections had
///    errors — so notebook success is not evidence that a section recovered.
///  • A section sync success covers that section's sync failures, its real-time channel failures, and
///    page-level failures known to belong to it. A full section sync completing IS evidence for those.
///  • A page upload/download success covers NOTHING. One page reaching the server says nothing about
///    the rest of the section. (The previous build cleared a whole section's failure on a single page
///    upload — that is the bug this type exists to prevent.)
///  • An unrecognised/suspected event covers nothing and is covered by nothing; it clears only via TTL,
///    because we cannot know what would prove it resolved.
/// </summary>
public static class SyncScope
{
    public const string NotebookPrefix = "notebook:";
    public const string SectionPrefix = "section:";
    public const string RealTimePrefix = "section-rt:";
    public const string PagePrefix = "page:";
    public const string EventPrefix = "event:";

    /// <summary>The scope an issue from this event belongs to, plus the section it sits in (when known).</summary>
    public static (string Key, string? SectionRid) Of(SyncEvent e, ResolvedNames names)
    {
        var sectionRid = e.SectionResourceId ?? e.SectionGosid ?? names.Section;
        return e.Kind switch
        {
            SyncEventKind.SectionSyncResult => (SectionPrefix + (sectionRid ?? "?"), sectionRid),
            SyncEventKind.RealTimeService => (RealTimePrefix + (sectionRid ?? "?"), sectionRid),
            SyncEventKind.PageUpload or SyncEventKind.PageDownload or SyncEventKind.PageSyncSession
                => (PagePrefix + (e.PageGoid ?? sectionRid ?? "?"), sectionRid),
            SyncEventKind.UnrecognisedStorage or SyncEventKind.OtherOneNoteSignal
                => (EventPrefix + e.EventName, sectionRid),
            _ => (NotebookPrefix + (e.NotebookGosid ?? e.NotebookResourceId ?? names.Notebook ?? "?"), sectionRid),
        };
    }

    /// <summary>
    /// The scope keys a SUCCESSFUL event proves healthy. Empty means it proves nothing about any
    /// outstanding error (the fail-closed default).
    /// </summary>
    public static IReadOnlyList<string> CoveredBy(SyncEvent e, ResolvedNames names)
    {
        if (e.Outcome != SyncOutcome.Success) return [];
        var sectionRid = e.SectionResourceId ?? e.SectionGosid ?? names.Section;
        return e.Kind switch
        {
            // a completed section sync covers that section's sync + real-time failures
            SyncEventKind.SectionSyncResult when sectionRid is not null =>
                [SectionPrefix + sectionRid, RealTimePrefix + sectionRid],
            // a real-time session that reported "No error" covers only the real-time channel
            SyncEventKind.RealTimeService when sectionRid is not null => [RealTimePrefix + sectionRid],
            // a notebook sync covers the notebook only — never its sections (see the type remarks)
            SyncEventKind.NotebookSyncResult =>
                [NotebookPrefix + (e.NotebookGosid ?? e.NotebookResourceId ?? names.Notebook ?? "?")],
            // one page reaching the server proves nothing about the section
            _ => [],
        };
    }

    /// <summary>True when a section-sync success for <paramref name="sectionRid"/> also covers this issue's page.</summary>
    public static bool SectionCoversPage(string scopeKey, string? issueSectionRid, string sectionRid) =>
        scopeKey.StartsWith(PagePrefix, StringComparison.Ordinal)
        && issueSectionRid is not null
        && string.Equals(issueSectionRid, sectionRid, StringComparison.OrdinalIgnoreCase);
}
