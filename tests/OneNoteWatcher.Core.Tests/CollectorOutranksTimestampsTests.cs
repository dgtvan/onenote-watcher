using OneNoteWatcher.Core.Detection;
using OneNoteWatcher.Core.Graph;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Status;

namespace OneNoteWatcher.Core.Tests;

public class SectionKeyTests
{
    [Theory]
    // every spelling of the same section seen in the wild must reduce to one key
    [InlineData("36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2")]          // sync event resource id
    [InlineData("0-36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2")]        // Graph id
    [InlineData("0|36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2")]        // FileIdentifier
    [InlineData("8d499fcb43aa42bd9a75d6d555478ac2")]                            // bare token
    public void All_spellings_reduce_to_the_same_key(string id)
    {
        Assert.Equal("8d499fcb43aa42bd9a75d6d555478ac2", SectionKey.Normalize(id));
    }

    [Fact]
    public void Different_sections_do_not_collide()
    {
        Assert.False(SectionKey.Same(
            "36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2",
            "0-36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("{F871B437-99B3-4241-A3F4-FFCBDE99566E}{1}")]   // a GOSID is not a section resource id
    public void Unmatchable_ids_never_compare_equal(string? id)
    {
        Assert.Null(SectionKey.Normalize(id));
        Assert.False(SectionKey.Same(id, id));   // unknown must never satisfy a match
    }
}

/// <summary>
/// Reported 2026-09-06: "Van / Quick Notes — a local change at 15:28 has not reached OneDrive after
/// 22 min", while OneNote's own sync dialog said "Up to date" and the collector had recorded a full
/// section sync of that section at 15:28:45. Both inputs this detector infers from can mislead — Graph
/// lags, and the local index re-stamps a merely re-synced section — so OneNote's own result wins.
/// </summary>
public class CollectorOutranksTimestampsTests
{
    private const string GraphId = "0-36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213";
    private const string EventRid = "36B934175DC7E3A4!s68429cef5b3a4a6587e5d7783bc7f213";

    private static GraphSnapshot Server(DateTimeOffset lm) => new(
        DateTimeOffset.UtcNow, [new GraphNotebook("nb", "Note", lm)],
        [new GraphSection(GraphId, "Quick Notes", lm, "nb", "Note", null, null)]);

    private static IReadOnlyList<LocalSection> Local(DateTimeOffset newest) =>
        [new LocalSection("Van", "Quick Notes", newest)];   // local index uses the notebook's nickname

    private static IReadOnlyList<SectionSyncState> Collector(DateTimeOffset at) =>
        [new SectionSyncState(SectionKey.Normalize(EventRid)!, "Note / Quick Notes", at)];

    private readonly DateTimeOffset _t0 = new(2026, 9, 5, 15, 17, 53, TimeSpan.Zero);
    private readonly DateTimeOffset _serverStale = new(2026, 9, 6, 8, 16, 25, TimeSpan.Zero);
    private readonly DateTimeOffset _localStamp = new(2026, 9, 6, 8, 28, 44, TimeSpan.Zero);
    private readonly DateTimeOffset _sectionSynced = new(2026, 9, 6, 8, 28, 45, TimeSpan.Zero);

    private OutcomeDetector Baselined()
    {
        var d = new OutcomeDetector(TimeSpan.FromMinutes(10));
        d.Evaluate(Local(_t0), Server(_serverStale), true, true, _t0);
        return d;
    }

    /// <summary>Poll once to start the "ahead of the server" clock, then again past the grace period —
    /// the alert needs both, which is exactly how the reported one arose.</summary>
    private IReadOnlyList<SyncIssue> AfterGrace(OutcomeDetector d, IReadOnlyList<SectionSyncState>? collector)
    {
        d.Evaluate(Local(_localStamp), Server(_serverStale), true, true, _localStamp, null);
        return d.Evaluate(Local(_localStamp), Server(_serverStale), true, true, _localStamp.AddMinutes(22), collector);
    }

    [Fact]
    public void The_exact_reported_case_raises_no_alert()
    {
        Assert.Empty(AfterGrace(Baselined(), Collector(_sectionSynced)));
    }

    [Fact]
    public void Without_the_collector_result_the_same_input_still_alerts()
    {
        var i = Assert.Single(AfterGrace(Baselined(), null));
        Assert.Equal(IssueKind.UploadStuck, i.Kind);
        Assert.Equal(GraphId, i.SectionId);   // tagged so it can be matched back to a collector result
    }

    [Fact]
    public void A_sync_well_before_the_local_change_proves_nothing()
    {
        Assert.Single(AfterGrace(Baselined(), Collector(_localStamp.AddMinutes(-5))));
    }

    [Fact]
    public void A_result_for_a_different_section_proves_nothing()
    {
        var other = new SectionSyncState(SectionKey.Normalize("36B934175DC7E3A4!s8d499fcb43aa42bd9a75d6d555478ac2")!,
            "Note / eSim Data", _sectionSynced);
        Assert.Single(AfterGrace(Baselined(), [other]));
    }

    [Fact]
    public void Once_confirmed_the_section_is_rebaselined_so_it_does_not_alert_later()
    {
        var d = Baselined();
        Assert.Empty(AfterGrace(d, Collector(_sectionSynced)));

        // hours later, still nothing new locally and the collector has since restarted (no results)
        var issues = d.Evaluate(Local(_localStamp), Server(_serverStale), true, true, _localStamp.AddHours(3), null);

        Assert.Empty(issues);
    }

    [Fact]
    public void An_ambiguous_section_name_is_not_matched_to_the_wrong_notebook()
    {
        // OneNote creates a "Quick Notes" in every notebook; the local notebook name ("Van") does not
        // match Graph's ("Note"), so a name-only fallback could compare two unrelated sections.
        var d = new OutcomeDetector(TimeSpan.FromMinutes(10));
        IReadOnlyList<LocalSection> two =
        [
            new LocalSection("Van", "Quick Notes", _t0),
            new LocalSection("Work", "Quick Notes", _localStamp),
        ];
        d.Evaluate(two, Server(_serverStale), true, true, _t0);
        var issues = d.Evaluate(two, Server(_serverStale), true, true, _localStamp.AddMinutes(22));

        Assert.Empty(issues);   // ambiguous → no comparison, rather than a comparison against the wrong section
    }

    // ---- giving up on a claim that can never be disproved ----

    /// <summary>Minutes of simulated polling elapsed, so successive calls never move the clock backwards.</summary>
    private double _elapsed;
    /// <summary>Everything the detector abandoned across the whole run — LastAbandoned covers one poll only.</summary>
    private readonly List<string> _abandoned = [];

    /// <summary>Keep polling every 5 minutes for a further <paramref name="hours"/>, returning the last result.</summary>
    private IReadOnlyList<SyncIssue> PollFor(OutcomeDetector d, double hours, bool collectorUp, bool oneNoteRunning)
    {
        IReadOnlyList<SectionSyncState>? sections = collectorUp ? [] : null;
        IReadOnlyList<SyncIssue> last = [];
        var until = _elapsed + hours * 60;
        for (; _elapsed <= until; _elapsed += 5)
        {
            last = d.Evaluate(Local(_localStamp), Server(_serverStale), oneNoteRunning, true,
                _localStamp.AddMinutes(_elapsed), sections);
            _abandoned.AddRange(d.LastAbandoned);
        }
        return last;
    }

    [Fact]
    public void An_uncorroborated_claim_is_abandoned_once_OneNote_and_the_collector_have_stayed_healthy()
    {
        var d = Baselined();
        Assert.NotEmpty(PollFor(d, 1, collectorUp: true, oneNoteRunning: true));    // still standing at 1 h

        var after = PollFor(d, 3, collectorUp: true, oneNoteRunning: true);         // past the 2 h default

        Assert.Empty(after);
    }

    [Fact]
    public void It_says_why_it_gave_up()
    {
        var d = Baselined();
        PollFor(d, 3, collectorUp: true, oneNoteRunning: true);

        var note = Assert.Single(_abandoned);
        Assert.Contains("Quick Notes", note);
        Assert.Contains("no corroboration", note);
    }

    [Fact]
    public void It_never_gives_up_while_OneNote_is_closed()
    {
        // "edited, then closed OneNote" is exactly what this check exists to catch — it must not expire
        var d = Baselined();
        Assert.NotEmpty(PollFor(d, 6, collectorUp: true, oneNoteRunning: false));
    }

    [Fact]
    public void It_never_gives_up_while_the_collector_is_down()
    {
        // without the collector a real failure would go unseen, so silence is not evidence of health
        var d = Baselined();
        Assert.NotEmpty(PollFor(d, 6, collectorUp: false, oneNoteRunning: true));
    }

    [Fact]
    public void The_give_up_rule_can_be_switched_off()
    {
        var d = new OutcomeDetector(TimeSpan.FromMinutes(10), TimeSpan.Zero);
        d.Evaluate(Local(_t0), Server(_serverStale), true, true, _t0);
        Assert.NotEmpty(PollFor(d, 12, collectorUp: true, oneNoteRunning: true));
    }
}
