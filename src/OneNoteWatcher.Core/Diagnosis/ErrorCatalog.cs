using System.Text.RegularExpressions;

namespace OneNoteWatcher.Core.Diagnosis;

/// <summary>Root-cause category of a sync failure — drives the recommendation and the ignore/transient policy.</summary>
public enum FailureCategory
{
    Unknown,
    Network,        // no internet / cannot reach OneDrive service / timeouts
    Service,        // server-side error, throttling, outage (0xE4xxxxxx family)
    Permission,     // sign-in, access denied, unauthorized, licensing
    Storage,        // OneDrive quota / disk space
    Password,       // password-protected section can't sync
    Corruption,     // section/notebook file damaged
    Conflict,       // merge/version conflicts
    FileState,      // file being renamed/moved/locked on the server
    ClientVersion,  // OneNote too old for the service
    Transient,      // busy/retry-later; resolves itself
}

/// <summary>What happened, why, and what to do about it — in plain language.</summary>
public sealed record ErrorExplanation(
    FailureCategory Category,
    string Summary,
    string Recommendation,
    bool LikelyTransient = false,
    string? Reference = null)
{
    public const string MicrosoftFixArticle =
        "https://support.microsoft.com/office/fix-issues-when-you-can-t-sync-onenote";
}

/// <summary>
/// Maps OneNote sync error codes / symbolic descriptions to human explanations and concrete fixes.
/// Codes are from Microsoft's "Fix issues when you can't sync OneNote" article plus values observed
/// on this machine (docs/reference-data-sources.md). Unknown codes still get a category from the
/// description keywords and the code family, so the user is never left with a bare number.
/// </summary>
public static class ErrorCatalog
{
    private static readonly Dictionary<uint, ErrorExplanation> ByCode = new()
    {
        [0xE000002E] = new(FailureCategory.Conflict,
            "The notebook or section is out of sync with the copy in OneDrive.",
            "In OneNote press Shift+F9 (Sync This Notebook Now). If it keeps failing, close the notebook " +
            "(right-click → Close This Notebook) and reopen it from OneDrive. As a last resort copy the " +
            "affected pages into a new section.",
            Reference: ErrorExplanation.MicrosoftFixArticle),

        [0xE000005D] = new(FailureCategory.FileState,
            "The section file is pending a rename/move on OneDrive, so OneNote cannot write to it right now.",
            "Usually clears on the next sync. If it persists for more than ~15 minutes, close and reopen the " +
            "notebook; check OneDrive online for a stuck rename of that section.",
            LikelyTransient: true),

        [0xE000005E] = new(FailureCategory.Conflict,
            "A section refers to a page revision that no longer exists on the server (ReferencedRevisionNotFound).",
            "Sync manually (Shift+F9). If the section still won't sync, right-click it → Close, then reopen " +
            "the notebook; copy the pages to a new section if the error returns.",
            Reference: ErrorExplanation.MicrosoftFixArticle),

        [0xE000012E] = new(FailureCategory.Corruption,
            "The local section file is damaged (FileNodeFileCorrupt).",
            "Open File → Info → Open Backups and restore the section, or copy the pages you can still open " +
            "into a new section and delete the damaged one. Check Notebook Recycle Bin / Misplaced Sections."),

        [0xE0000320] = new(FailureCategory.Password,
            "A password-protected section is locked, so OneNote skips it (its already-synced notes are unaffected). " +
            "Only changes made while locked stay on this PC until you unlock it.",
            "Nothing to do unless you edited that section: unlock it (click it, enter the password) and it syncs. " +
            "If the password was changed on another device, unlock it here with the new one.",
            LikelyTransient: true),

        [0xE000078B] = new(FailureCategory.Network,
            "The server timed out while syncing.",
            "Check your internet connection and any VPN/proxy; wait a few minutes and sync again (Shift+F9).",
            LikelyTransient: true),

        [0xE0000796] = new(FailureCategory.Storage,
            "Your OneDrive storage quota is exceeded.",
            "Free up space in OneDrive (empty its recycle bin, delete large files) or upgrade storage; sync resumes automatically."),

        [0xE00009C8] = new(FailureCategory.Service,
            "Sync was interrupted by a service-side issue (acknowledged by the OneNote team).",
            "Make sure OneNote is up to date (File → Account → Update Options → Update Now), then sync manually.",
            LikelyTransient: true, Reference: ErrorExplanation.MicrosoftFixArticle),

        [0xE00015E0] = new(FailureCategory.Storage,
            "Not enough space on OneDrive to store the changes.",
            "Free up OneDrive storage, then sync again."),

        [0xE4010641] = new(FailureCategory.Network,
            "OneNote cannot reach the OneDrive service.",
            "Check internet connectivity, VPN and proxy settings; try opening onedrive.live.com in a browser. " +
            "If the web works but OneNote doesn't, sign out and back in (File → Account)."),

        [0xE4010690] = new(FailureCategory.Network,
            "OneNote cannot reach the server.",
            "Check connectivity/VPN/proxy; verify the Microsoft 365 service status; retry in a few minutes."),

        [0xE40105F9] = new(FailureCategory.ClientVersion,
            "This version of OneNote is no longer supported by the service.",
            "Update Office: File → Account → Update Options → Update Now, then restart OneNote."),

        [0xE4020040] = new(FailureCategory.Transient,
            "The OneDrive store is busy.",
            "No action needed unless it persists; OneNote retries automatically.",
            LikelyTransient: true),
    };

    // description-keyword fallbacks, checked in order
    private static readonly (Regex Pattern, ErrorExplanation Explanation)[] ByDescription =
    [
        (Rx("Crypto|Passphrase|Password"), new(FailureCategory.Password,
            "A password-protected section cannot sync while locked or with a mismatched passphrase.",
            "Unlock the section with its password, then sync.")),
        (Rx("Corrupt|Malformed|RanOffEnd|Checksum"), new(FailureCategory.Corruption,
            "A section or notebook file appears damaged.",
            "Restore from File → Info → Open Backups, or copy readable pages to a new section.")),
        (Rx("Quota|Space|Storage|Full"), new(FailureCategory.Storage,
            "OneDrive or local storage is full.",
            "Free up space in OneDrive / on disk, then sync.")),
        (Rx("Unauthori|Forbidden|Access|Permission|Denied|Auth|Token|Login|Sign"), new(FailureCategory.Permission,
            "OneNote is not allowed to access the notebook location.",
            "Sign out and back in (File → Account), and confirm you still have access to the notebook in OneDrive. " +
            "For a shared notebook, ask the owner to re-share it.")),
        (Rx("Timeout|Timed|Network|Connect|Http|Dns|Offline|Unreachable|Socket"), new(FailureCategory.Network,
            "The network or the OneDrive service could not be reached.",
            "Check Wi-Fi/VPN/proxy; verify onedrive.live.com opens; retry.", LikelyTransient: true)),
        (Rx("Conflict|Merge|Revision|OutOfSync"), new(FailureCategory.Conflict,
            "Local and server copies conflict.",
            "Sync manually; look for 'conflict' pages in the section and merge them; close and reopen the notebook if needed.")),
        (Rx("PendingRename|Rename|Moved|Locked|InUse|Lock"), new(FailureCategory.FileState,
            "The section file is being renamed/moved or is locked on the server.",
            "Wait for the next sync; if it persists, close and reopen the notebook.", LikelyTransient: true)),
        (Rx("Busy|Throttl|Retry|Later|TryAgain"), new(FailureCategory.Transient,
            "The service asked OneNote to retry later.",
            "No action needed unless it persists.", LikelyTransient: true)),
    ];

    public static ErrorExplanation Explain(uint code, string? description, string? errorType = null)
    {
        if (code != 0 && ByCode.TryGetValue(code, out var known)) return known;

        if (!string.IsNullOrWhiteSpace(description))
            foreach (var (rx, ex) in ByDescription)
                if (rx.IsMatch(description)) return ex with
                {
                    Summary = $"{ex.Summary} ({description})",
                };

        // family fallback: 0xE4xxxxxx = service/network side, 0xE0xxxxxx = local store
        var family = (code & 0xFF000000u) switch
        {
            0xE4000000u => FailureCategory.Service,
            0xE0000000u => FailureCategory.Unknown,
            _ => FailureCategory.Unknown,
        };
        var codeText = code == 0 ? "" : $" 0x{code:X8}";
        return new ErrorExplanation(
            family,
            $"OneNote reported an unrecognised sync error{codeText}" +
            (description is null ? "." : $" ({description})."),
            "Sync manually (Shift+F9). If it persists, look the code up in Microsoft's article and check " +
            "File → Info → View Sync Status in OneNote for the message shown there.",
            Reference: ErrorExplanation.MicrosoftFixArticle);
    }

    /// <summary>Explanation for an outcome-based finding (no code): content not reaching the cloud.</summary>
    public static ErrorExplanation UploadStuck(bool oneNoteRunning, bool? internetAvailable) =>
        internetAvailable == false
            ? new(FailureCategory.Network,
                "Local changes are not reaching OneDrive because the PC is offline.",
                "Reconnect to the internet; OneNote will upload automatically.", LikelyTransient: true)
            : !oneNoteRunning
            ? new(FailureCategory.Unknown,
                "Local changes exist that were not uploaded before OneNote was closed.",
                "Start OneNote and let it sync; keep it open until the section shows as synced.")
            : new(FailureCategory.Unknown,
                "OneNote is running and online, but local changes are not reaching OneDrive.",
                "Press Shift+F9 to force a sync and open File → Info → View Sync Status to see the error OneNote reports; " +
                "the collector's history will carry the exact code as soon as OneNote emits it.");

    private static Regex Rx(string p) => new(p, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
}
