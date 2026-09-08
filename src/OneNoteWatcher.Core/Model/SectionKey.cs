using System.Text.RegularExpressions;

namespace OneNoteWatcher.Core.Model;

/// <summary>
/// One canonical key for a section, whatever spelling a source used.
///
/// The same section arrives under at least four spellings (see docs/reference-data-sources.md):
/// the sync-event resource id <c>36B934175DC7E3A4!s8d49…</c>, Graph's id <c>0-36B934175DC7E3A4!s8d49…</c>,
/// the <c>0|…</c> FileIdentifier form, and the bare <c>8d49…</c> token. Matching a collector result to a
/// cloud-check finding fails silently unless they are reduced to the same key first — and a silent
/// match failure means an alert that can never be answered.
/// </summary>
public static class SectionKey
{
    // the section token in "…!s<hex>" (Graph, sync events, FileIdentifier), else a standalone hex run
    private static readonly Regex AfterBang = new(@"![sS]([0-9a-fA-F]{16,})", RegexOptions.Compiled);
    private static readonly Regex BareToken = new(@"(?<![0-9a-fA-F])([0-9a-fA-F]{24,})(?![0-9a-fA-F])", RegexOptions.Compiled);

    /// <summary>
    /// A fifth spelling: a section whose OneDrive item id is a short number rather than an "!s" token —
    /// <c>36B934175DC7E3A4!1242</c> in sync events, <c>0-36B934175DC7E3A4!1242</c> in Graph. Neither rule
    /// above can key it: there is no "!s", and the drive id is 16 hex where <see cref="BareToken"/> needs
    /// 24+. Every such id normalised to null, and a null key is not a near miss — it silently discards the
    /// collector's proof that the section synced and leaves the cloud check's "not reaching OneDrive"
    /// claim with nothing able to answer it. Observed 2026-09-08: 16 of this machine's 72 sections use
    /// this shape, and "Van / Family" was reported stuck for 1.2 h while the sync history for that exact
    /// second showed it syncing healthily.
    ///
    /// The key must be BOTH parts. The drive id alone is shared by every section in the drive, so keying
    /// on it would merge unrelated sections — trading a missed match for a wrong one.
    /// </summary>
    private static readonly Regex DriveItem = new(@"([0-9a-fA-F]{8,})!([0-9a-zA-Z]+)", RegexOptions.Compiled);

    /// <summary>The canonical key, or null when the id carries no section token we can match on.</summary>
    public static string? Normalize(string? id)
    {
        if (string.IsNullOrWhiteSpace(id)) return null;
        var m = AfterBang.Match(id);
        if (m.Success) return m.Groups[1].Value.ToLowerInvariant();
        m = BareToken.Match(id);
        if (m.Success) return m.Groups[1].Value.ToLowerInvariant();
        // last, so it only ever catches ids that previously produced null — it can widen matching, never
        // re-interpret an id the rules above already key
        m = DriveItem.Match(id);
        return m.Success ? $"{m.Groups[1].Value}!{m.Groups[2].Value}".ToLowerInvariant() : null;
    }

    /// <summary>True when both ids denote the same section. False whenever either cannot be normalised.</summary>
    public static bool Same(string? a, string? b)
    {
        var x = Normalize(a);
        return x is not null && x == Normalize(b);
    }
}
