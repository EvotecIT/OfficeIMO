using System.Text;
using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

public class YamlFormatterTests {
    public static TheoryData<string> Layouts => new() { "default", "indent4-indented", "compact-double-crlf", "indent3-single-nofinal" };

    private static YamlFormatOptions Layout(string name) => name switch {
        "indent4-indented" => new YamlFormatOptions { IndentWidth = 4, SequenceIndentation = YamlSequenceIndentation.Indented },
        "compact-double-crlf" => new YamlFormatOptions { SequenceIndentation = YamlSequenceIndentation.Compact, Quotes = YamlQuoteStyle.Double, LineEnding = YamlLineEnding.CrLf },
        "indent3-single-nofinal" => new YamlFormatOptions { IndentWidth = 3, Quotes = YamlQuoteStyle.Single, FinalNewline = false },
        _ => new YamlFormatOptions()
    };

    [Theory]
    [MemberData(nameof(Layouts))]
    public void FormattingTheCorpusKeepsValuesAndCommentsAndIsIdempotent(string layout) {
        YamlFormatOptions options = Layout(layout);
        int formatted = 0;
        foreach (YamlCorpusFile file in YamlCorpus.Load()) {
            string source = YamlTextTestHelpers.DecodeSuiteInput(file.Bytes);
            string once;
            try {
                once = YamlFormatter.Format(source, options);
            } catch (YamlException) {
                continue;
            }
            formatted++;
            string twice = YamlFormatter.Format(once, options);
            Assert.True(once == twice, file.Path + ": formatting again changed the text.");
            using YamlDocument before = YamlDocument.Parse(source);
            using YamlDocument after = YamlDocument.Parse(once);
            Assert.True(YamlValueEquivalence.SameValues(before, after), file.Path + ": formatting changed a value.");
            Assert.Equal(Comments(before), Comments(after));
            IReadOnlyList<YamlLintFinding> findings = YamlLinter.Lint(once, MatchingLint(options));
            Assert.True(findings.Count == 0, file.Path + ": " + string.Join("; ", findings));
        }
        Assert.True(formatted >= 17, "Only " + formatted + " corpus files were formatted.");
    }

    [Fact]
    public void FormattingTheSpecificationSuiteKeepsValuesAndIsIdempotent() {
        int formatted = 0;
        foreach (YamlTestSuiteCase test in YamlTestSuite.Load().Cases) {
            if (test.IsError) continue;
            string source = YamlTextTestHelpers.DecodeSuiteInput(test.Input);
            YamlDocument before;
            try {
                before = YamlDocument.Parse(source);
            } catch (YamlException) {
                continue;
            }
            using (before) {
                string once = YamlFormatter.Format(source);
                using YamlDocument after = YamlDocument.Parse(once);
                Assert.True(YamlValueEquivalence.SameValues(before, after), test.Id + ": formatting changed a value.");
                Assert.True(once == YamlFormatter.Format(once), test.Id + ": formatting again changed the text.");
                formatted++;
            }
        }
        Assert.True(formatted > 300, "Only " + formatted + " suite cases were formatted.");
    }

    [Fact]
    public void CommentAboveAnInlineRootStaysOnItsOwnLine() {
        Assert.Equal("# note\n!foo \"bar\"\n", YamlFormatter.Format("# note\n!foo \"bar\"\n"));
        Assert.Equal("---\n# note\n{a: 1}\n", YamlFormatter.Format("---\n# note\n{ a: 1 }\n"));
        Assert.Equal("---\n---\n", YamlFormatter.Format("---\n---\n"));
    }

    /// <summary>The structural rules the formatter promises to satisfy for its own layout.</summary>
    private static YamlLintOptions MatchingLint(YamlFormatOptions options) {
        var lint = new YamlLintOptions();
        foreach (YamlLintRuleOptions rule in lint.Rules) rule.Enabled = false;
        lint.Indentation.Enabled = true;
        lint.Indentation.Spaces = options.IndentWidth;
        lint.Indentation.Sequences = options.SequenceIndentation switch {
            YamlSequenceIndentation.Compact => YamlLintSequenceIndentation.NotIndented,
            YamlSequenceIndentation.Indented => YamlLintSequenceIndentation.Indented,
            _ => YamlLintSequenceIndentation.Consistent
        };
        lint.TrailingSpaces.Enabled = true;
        lint.Tabs.Enabled = true;
        return lint;
    }

    private static string Comments(YamlDocument document) {
        var text = new StringBuilder();
        foreach (YamlDocumentPart part in document.Documents) {
            text.Append(part.CommentBefore).Append('|').Append(part.Comment).Append('|').Append(part.CommentAfter).Append('\n');
            if (part.RootNode is null) continue;
            foreach (YamlNode node in part.RootNode.Descendants()) {
                if (node.CommentBefore is { } before) text.Append(before).Append('\n');
                if (node.Comment is { } comment) text.Append(comment).Append('\n');
            }
        }
        return text.ToString();
    }

    [Fact]
    public void IndentWidthAndSequenceStyleAreApplied() {
        string yaml = "a:\n   b: 1\n   list:\n   - x\n   -   y\n";
        Assert.Equal("a:\n    b: 1\n    list:\n        - x\n        - y\n",
            YamlFormatter.Format(yaml, new YamlFormatOptions { IndentWidth = 4, SequenceIndentation = YamlSequenceIndentation.Indented }));
        Assert.Equal("a:\n  b: 1\n  list:\n  - x\n  - y\n",
            YamlFormatter.Format(yaml, new YamlFormatOptions { IndentWidth = 2, SequenceIndentation = YamlSequenceIndentation.Compact }));
    }

    [Fact]
    public void QuoteStyleChangesQuotesButNotValues() {
        string yaml = "a: 'one'\nb: \"two\"\nc: \"bell\\a\"\nd: plain\ne: 'yes'\n";
        Assert.Equal("a: \"one\"\nb: \"two\"\nc: \"bell\\a\"\nd: plain\ne: \"yes\"\n", YamlFormatter.Format(yaml, new YamlFormatOptions { Quotes = YamlQuoteStyle.Double }));
        // The bell character needs a double-quoted escape, so that value keeps its double quotes.
        Assert.Equal("a: 'one'\nb: 'two'\nc: \"bell\\a\"\nd: plain\ne: 'yes'\n", YamlFormatter.Format(yaml, new YamlFormatOptions { Quotes = YamlQuoteStyle.Single }));
        Assert.Equal(yaml, YamlFormatter.Format(yaml));
    }

    [Fact]
    public void AnchorsTagsAndCommentsSurvive() {
        string yaml = "# header\n\nbase: &base\n    port: 80   # http\nweb:\n    <<: *base\n    when: !!str 2024-01-01\n    run: !Ref name\n";
        Assert.Equal("# header\n\nbase: &base\n  port: 80 # http\nweb:\n  <<: *base\n  when: !!str 2024-01-01\n  run: !Ref name\n",
            YamlFormatter.Format(yaml, new YamlFormatOptions { IndentWidth = 2 }));
    }

    [Fact]
    public void LineEndingAndFinalNewlineAreApplied() {
        Assert.Equal("a: 1\r\nb: 2\r\n", YamlFormatter.Format("a: 1\nb: 2\n", new YamlFormatOptions { LineEnding = YamlLineEnding.CrLf }));
        Assert.Equal("a: 1\nb: 2\n", YamlFormatter.Format("a: 1\r\nb: 2\r\n", new YamlFormatOptions { LineEnding = YamlLineEnding.Lf }));
        Assert.Equal("a: 1\r\nb: 2\r\n", YamlFormatter.Format("a: 1\r\nb: 2\r\n"));
        Assert.Equal("a: 1\nb: 2", YamlFormatter.Format("a: 1\nb: 2\n", new YamlFormatOptions { FinalNewline = false }));
        // The line break is part of the block scalar's value, so it stays.
        Assert.Equal("a: |\n  text\n", YamlFormatter.Format("a: |\n  text\n", new YamlFormatOptions { FinalNewline = false }));
    }

    [Fact]
    public void FormattingAnEditedDocumentIncludesTheEdits() {
        using YamlDocument document = YamlDocument.Parse("a:\n    b: 1\n");
        document.Set("a.c", "new");
        Assert.Equal("a:\n  b: 1\n  c: new\n", YamlFormatter.Format(document, new YamlFormatOptions { IndentWidth = 2 }));
        Assert.Equal("a:\n    b: 1\n    c: new\n", document.ToString());
    }

    [Fact]
    public void InvalidYamlThrows() {
        Assert.Throws<YamlException>(() => YamlFormatter.Format("a: [1, 2\n"));
    }
}
