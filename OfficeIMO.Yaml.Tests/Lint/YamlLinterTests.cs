using Xunit;

namespace OfficeIMO.Yaml.Tests;

public class YamlLinterTests {
    private static YamlLintFinding Single(string yaml, string rule, YamlLintOptions? options = null) =>
        Assert.Single(YamlLinter.Lint(yaml, options), f => f.RuleId == rule);

    private static void None(string yaml, string rule, YamlLintOptions? options = null) =>
        Assert.DoesNotContain(YamlLinter.Lint(yaml, options), f => f.RuleId == rule);

    [Fact]
    public void CleanFileHasNoFindings() {
        Assert.Empty(YamlLinter.Lint("# settings\nname: web\nports:\n- 80\n- 443\nlabels:\n  app: web  # inline\n"));
    }

    [Fact]
    public void SyntaxErrorIsReportedWithLineRulesButNoFixes() {
        IReadOnlyList<YamlLintFinding> findings = YamlLinter.Lint("a: 1   \nb: [1, 2\n");
        YamlLintFinding syntax = Assert.Single(findings, f => f.RuleId == YamlLintRules.Syntax);
        Assert.Equal(OfficeDiagnosticSeverity.Error, syntax.Severity);
        YamlLintFinding spaces = Assert.Single(findings, f => f.RuleId == YamlLintRules.TrailingSpaces);
        Assert.Null(spaces.Fix);
    }

    [Fact]
    public void TrailingSpacesAreFixedOutsideScalarsAndOnlyReportedInsideBlockScalars() {
        string yaml = "a: 1  \nscript: |\n  echo hi  \n  done\nb: 2\t\n";
        YamlLintFinding[] findings = YamlLinter.Lint(yaml).Where(f => f.RuleId == YamlLintRules.TrailingSpaces).ToArray();
        Assert.Equal(new[] { 1, 3, 5 }, findings.Select(f => f.Line));
        Assert.Equal(new[] { 5, 10, 5 }, findings.Select(f => f.Column));
        Assert.NotNull(findings[0].Fix);
        Assert.Null(findings[1].Fix);
        Assert.NotNull(findings[2].Fix);
    }

    [Fact]
    public void LineLengthAllowsOneLongWordAndKeyWithOneLongWord() {
        var options = new YamlLintOptions { LineLength = { Max = 20 } };
        YamlLintFinding finding = Single("text: one two three four five six\n", YamlLintRules.LineLength, options);
        Assert.Equal(21, finding.Column);
        None("url: https://example.com/a/very/long/path\n", YamlLintRules.LineLength, options);
        None("- https://example.com/a/very/long/path\n", YamlLintRules.LineLength, options);
        options.LineLength.AllowNonBreakableInlineMappings = false;
        Single("url: https://example.com/a/very/long/path\n", YamlLintRules.LineLength, options);
    }

    [Fact]
    public void TruthyReportsYaml11BooleansAndQuotesThemUnderTheCoreSchema() {
        YamlLintFinding finding = Single("enabled: yes\non: push\n", YamlLintRules.Truthy);
        Assert.Equal(1, finding.Line);
        Assert.Equal("'yes'", finding.Fix!.Replacement);
        None("enabled: true\nother: 'yes'\nflag: TRUE\n", YamlLintRules.Truthy);
        Single("on: push\n", YamlLintRules.Truthy, new YamlLintOptions { Truthy = { CheckKeys = true } });
    }

    [Fact]
    public void TruthyRewritesToTrueOrFalseWhereTheValueIsABoolean() {
        // Under %YAML 1.1 'off' is a boolean, so the fix spells the same boolean the 1.2 way.
        YamlLintFinding finding = Single("%YAML 1.1\n---\nenabled: off\n", YamlLintRules.Truthy);
        Assert.Equal("false", finding.Fix!.Replacement);
        var lowercase = new YamlLintOptions();
        lowercase.Truthy.AllowedValues.Remove("True");
        Assert.Equal("true", Single("a: True\n", YamlLintRules.Truthy, lowercase).Fix!.Replacement);
    }

    [Fact]
    public void KeyDuplicatesIsAnErrorAtTheRepeatedKey() {
        YamlLintFinding finding = Single("a: 1\nb: 2\na: 3\n", YamlLintRules.KeyDuplicates);
        Assert.Equal(OfficeDiagnosticSeverity.Error, finding.Severity);
        Assert.Equal(3, finding.Line);
        Assert.Null(finding.Fix);
        None("a: 1\n\"1\": 2\n1: 3\n", YamlLintRules.KeyDuplicates);
    }

    [Fact]
    public void KeyDuplicatesUnderAStrictLoadIsStillAKeyDuplicatesFinding() {
        var options = new YamlLintOptions { LoadOptions = new YamlLoadOptions { Strict = true } };
        IReadOnlyList<YamlLintFinding> findings = YamlLinter.Lint("a: 1\na: 2\n", options);
        Assert.Contains(findings, f => f.RuleId == YamlLintRules.KeyDuplicates && f.Line == 2);
        Assert.DoesNotContain(findings, f => f.RuleId == YamlLintRules.Syntax);
    }

    [Fact]
    public void IndentationFollowsTheFirstWidthUnlessConfigured() {
        string yaml = "a:\n  b:\n    c: 1\nd:\n    e: 1\n";
        YamlLintFinding finding = Single(yaml, YamlLintRules.Indentation);
        Assert.Equal(5, finding.Line);
        Assert.Contains("expected 2 but found 4", finding.Message, StringComparison.Ordinal);
        Assert.Equal(new[] { 2, 3 }, YamlLinter.Lint(yaml, new YamlLintOptions { Indentation = { Spaces = 4 } })
            .Where(f => f.RuleId == YamlLintRules.Indentation).Select(f => f.Line));
    }

    [Fact]
    public void SequenceIndentationIsConsistentByDefault() {
        YamlLintFinding finding = Single("a:\n- 1\nb:\n  - 2\n", YamlLintRules.Indentation);
        Assert.Equal(4, finding.Line);
        None("a:\n  - 1\nb:\n  - 2\n", YamlLintRules.Indentation);
        Single("a:\n  - 1\n", YamlLintRules.Indentation, new YamlLintOptions { Indentation = { Sequences = YamlLintSequenceIndentation.NotIndented } });
        None("a:\n- 1\nb:\n  - 2\n", YamlLintRules.Indentation, new YamlLintOptions { Indentation = { Sequences = YamlLintSequenceIndentation.Any } });
    }

    [Fact]
    public void DocumentStartIsOffByDefaultAndCanBeRequiredOrForbidden() {
        None("a: 1\n", YamlLintRules.DocumentStart);
        var required = new YamlLintOptions { DocumentStart = { Enabled = true } };
        YamlLintFix fix = Single("a: 1\n", YamlLintRules.DocumentStart, required).Fix!;
        Assert.Equal((0, 0, "---\n"), (fix.Offset, fix.Length, fix.Replacement));
        None("---\na: 1\n", YamlLintRules.DocumentStart, required);
        var forbidden = new YamlLintOptions { DocumentStart = { Enabled = true, Present = false } };
        Assert.NotNull(Single("---\na: 1\n", YamlLintRules.DocumentStart, forbidden).Fix);
        // A comment on the marker line would be lost, so the marker is reported but kept.
        Assert.Null(Single("--- # generated\na: 1\n", YamlLintRules.DocumentStart, forbidden).Fix);
        None("%YAML 1.2\n---\na: 1\n", YamlLintRules.DocumentStart, forbidden);
    }

    [Fact]
    public void EmptyValuesAreReportedWhenEnabled() {
        string yaml = "a:\nb: 1\nc: {d: }\nitems:\n-\n- 1\n";
        None(yaml, YamlLintRules.EmptyValues);
        YamlLintFinding[] findings = YamlLinter.Lint(yaml, new YamlLintOptions { EmptyValues = { Enabled = true } })
            .Where(f => f.RuleId == YamlLintRules.EmptyValues).ToArray();
        Assert.Equal(new[] { 1, 3, 5 }, findings.Select(f => f.Line));
    }

    [Fact]
    public void CommentsNeedAStartingSpaceAndDistanceFromContent() {
        YamlLintFinding[] findings = YamlLinter.Lint("#!/usr/bin/env yq\n#comment\n######\na: 1 # close\nb: 'x #y'  # ok\n")
            .Where(f => f.RuleId == YamlLintRules.Comments).ToArray();
        Assert.Equal(new[] { (2, 2), (4, 6) }, findings.Select(f => (f.Line, f.Column)));
        Assert.Equal(" ", findings[0].Fix!.Replacement);
        Assert.Equal(" ", findings[1].Fix!.Replacement);
    }

    [Fact]
    public void OctalValuesReportImplicitOctalsByDefault() {
        YamlLintFinding finding = Single("mode: 0755\nport: 8080\nquoted: '0755'\nzero: 0\n", YamlLintRules.OctalValues);
        Assert.Equal(1, finding.Line);
        None("mode: 0o755\n", YamlLintRules.OctalValues);
        Single("mode: 0o755\n", YamlLintRules.OctalValues, new YamlLintOptions { OctalValues = { ForbidExplicitOctal = true } });
    }

    [Fact]
    public void TabsInIndentationAreReportedButTabsInsideScalarsAreNot() {
        YamlLintFinding finding = Single("a:\n\tb: 1\n", YamlLintRules.Tabs);
        Assert.Equal(2, finding.Line);
        None("a:\n  b: \"x\n   \ty\"\n", YamlLintRules.Tabs);
        None("b: |\n  \tindented\n", YamlLintRules.Tabs);
    }

    [Fact]
    public void FinalNewLineIsAddedUnlessABlockScalarEndsTheFile() {
        Assert.Equal("\n", Single("a: 1", YamlLintRules.NewLineAtEndOfFile).Fix!.Replacement);
        Assert.Equal("\r\n", Single("a: 1\r\nb: 2", YamlLintRules.NewLineAtEndOfFile).Fix!.Replacement);
        Assert.Null(Single("a: |\n  text", YamlLintRules.NewLineAtEndOfFile).Fix);
    }

    [Fact]
    public void EmptyLinesAreLimitedAtStartInsideAndEnd() {
        YamlLintFinding[] findings = YamlLinter.Lint("\na: 1\n\n\n\nb: |\n  x\n\n\n\n  y\nc: 2\n\n")
            .Where(f => f.RuleId == YamlLintRules.EmptyLines).ToArray();
        // Lines 8 to 10 are inside the block scalar and are its content.
        Assert.Equal(new[] { 1, 5, 13 }, findings.Select(f => f.Line));
    }

    [Fact]
    public void QuotedStringsChecksTypeAndRequirement() {
        var options = new YamlLintOptions { QuotedStrings = { Enabled = true, QuoteType = YamlLintQuoteType.Single } };
        YamlLintFinding[] findings = YamlLinter.Lint("a: plain\nb: \"double\"\nc: 'single'\nd: \"bell\\a\"\ne: 1\n", options)
            .Where(f => f.RuleId == YamlLintRules.QuotedStrings).ToArray();
        Assert.Equal(new[] { 1, 2 }, findings.Select(f => f.Line));
        Assert.Equal("'plain'", findings[0].Fix!.Replacement);
        Assert.Equal("'double'", findings[1].Fix!.Replacement);

        options.QuotedStrings.Required = YamlLintQuoteRequirement.OnlyWhenNeeded;
        findings = YamlLinter.Lint("a: plain\nb: 'needless'\nc: 'yes'\nd: '1.0'\ne: ': x'\n", options)
            .Where(f => f.RuleId == YamlLintRules.QuotedStrings).ToArray();
        YamlLintFinding redundant = Assert.Single(findings);
        Assert.Equal(2, redundant.Line);
        Assert.Equal("needless", redundant.Fix!.Replacement);
    }

    [Fact]
    public void KeyOrderingIsReportedWhenEnabled() {
        None("b: 1\na: 2\n", YamlLintRules.KeyOrdering);
        YamlLintFinding[] findings = YamlLinter.Lint("b: 1\na: 2\nc: {z: 1, y: 2}\n", new YamlLintOptions { KeyOrdering = { Enabled = true } })
            .Where(f => f.RuleId == YamlLintRules.KeyOrdering).ToArray();
        Assert.Equal(new[] { (2, 1), (3, 11) }, findings.Select(f => (f.Line, f.Column)));
    }

    [Fact]
    public void RulesCanBeDisabledAndChangeSeverity() {
        var options = new YamlLintOptions();
        options.GetRule("trailing-spaces")!.Enabled = false;
        options.GetRule(YamlLintRules.Truthy)!.Severity = OfficeDiagnosticSeverity.Error;
        IReadOnlyList<YamlLintFinding> findings = YamlLinter.Lint("a: yes  \n", options);
        YamlLintFinding finding = Assert.Single(findings);
        Assert.Equal((YamlLintRules.Truthy, OfficeDiagnosticSeverity.Error), (finding.RuleId, finding.Severity));
        Assert.Null(options.GetRule("colons"));
        Assert.StartsWith("1:4 error ", finding.ToString(), StringComparison.Ordinal);
        Assert.EndsWith(" (truthy)", finding.ToString(), StringComparison.Ordinal);
    }
}
