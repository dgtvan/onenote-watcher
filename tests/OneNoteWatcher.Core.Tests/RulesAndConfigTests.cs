using OneNoteWatcher.Core.Config;
using OneNoteWatcher.Core.Model;
using OneNoteWatcher.Core.Rules;

namespace OneNoteWatcher.Core.Tests;

public class PatternTests
{
    [Theory]
    [InlineData("0xE402*", "0xE4020040", true)]
    [InlineData("0xE402*", "0xE0000320", false)]
    [InlineData("Scratch*", "Scratch Notebook", true)]
    [InlineData("Work/Archive*", "Work/Archive 2024", true)]
    [InlineData("*store busy*", "The store busy, retry", true)]
    [InlineData("/are you sure/i", "Are You Sure you want to delete", true)]
    [InlineData("/^0xE000/i", "0xE000005D", true)]
    [InlineData("exact", "exactly", false)]
    public void Glob_and_regex(string pattern, string value, bool expected)
    {
        Assert.Equal(expected, Pattern.Parse(pattern).IsMatch(value));
    }
}

public class IgnoreRulesTests
{
    private static SyncIssue Issue(string? code = null, string msg = "",
        string? nb = null, string? sec = null, string detector = "etw") => new()
    {
        Detector = detector,
        Kind = IssueKind.ErrorCodeReported,
        Code = code, Message = msg, NotebookName = nb, SectionName = sec,
    };

    [Fact]
    public void Ignores_by_code_wildcard()
    {
        var rules = new IgnoreRules(["0xE402*"], [], [], [], []);
        Assert.True(rules.IsIgnored(Issue(code: "0xE4020040")));
        Assert.False(rules.IsIgnored(Issue(code: "0xE000005D")));
    }

    [Fact]
    public void Ignores_by_message_regex_notebook_section_and_detector()
    {
        var rules = new IgnoreRules(
            codes: [], messages: ["/store busy/i"], notebooks: ["Scratch*"],
            sections: ["Work/Archive*"], detectors: ["oalerts"]);

        Assert.True(rules.IsIgnored(Issue(msg: "The Store Busy now")));
        Assert.True(rules.IsIgnored(Issue(nb: "Scratch NB")));
        Assert.True(rules.IsIgnored(Issue(sec: "Work/Archive 2023")));
        Assert.True(rules.IsIgnored(Issue(detector: "oalerts")));
        Assert.False(rules.IsIgnored(Issue(code: "0xE000005D")));
    }

    [Fact]
    public void Empty_rules_ignore_nothing()
    {
        Assert.False(IgnoreRules.Empty.IsIgnored(Issue(code: "0xE000005D")));
    }
}

public class IniFileTests
{
    private const string Sample = """
        [general]
        poll_minutes = 5
        log_every_poll = true

        [graph]
        client_id = eecee773-5ffb-48f1-be33-41dff06da7f9   ; the app id
        tenant = consumers

        [ignore]
        codes     = 0xE4020040, 0xE402*
        messages  = /are you sure/i
        # a comment line
        """;

    [Fact]
    public void Reads_values_types_and_lists()
    {
        var ini = IniFile.Parse(Sample);
        Assert.Equal(5, ini.GetInt("general", "poll_minutes", 0));
        Assert.True(ini.GetBool("general", "log_every_poll", false));
        Assert.Equal("eecee773-5ffb-48f1-be33-41dff06da7f9", ini.Get("graph", "client_id"));
        Assert.Equal(["0xE4020040", "0xE402*"], ini.GetList("ignore", "codes"));
        Assert.Equal("consumers", ini.Get("graph", "tenant", "x"));
        Assert.Equal("fallback", ini.Get("nope", "missing", "fallback"));
    }

    [Fact]
    public void Strips_trailing_comment_but_keeps_value()
    {
        var ini = IniFile.Parse("[s]\nk = value ; comment\n");
        Assert.Equal("value", ini.Get("s", "k"));
    }

    [Fact]
    public void Builds_ignore_rules_from_the_shipped_config()
    {
        var ini = IniFile.Parse(Sample);
        var rules = IgnoreRules.FromConfig(ini);
        Assert.True(rules.IsIgnored(new SyncIssue
        {
            Detector = "etw", Kind = IssueKind.ErrorCodeReported, Code = "0xE4020040",
        }));
    }
}
