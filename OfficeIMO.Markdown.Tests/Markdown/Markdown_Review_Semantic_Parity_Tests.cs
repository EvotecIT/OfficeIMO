using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Review_Semantic_Parity_Tests {
    public static IEnumerable<object[]> FencedLineEndings() {
        for (int ending = 0; ending < 3; ending++) {
            for (int shape = 0; shape < 6; shape++) yield return new object[] { ending, shape };
        }
    }

    [Theory]
    [MemberData(nameof(FencedLineEndings))]
    public void Fenced_Code_Eof_Payload_Matches_Across_Reader_Modes(int ending, int shape) {
        string lineEnding = ending == 0 ? "\n" : ending == 1 ? "\r\n" : "\r";
        var (markdown, payload) = shape switch {
            0 => ("```\npayload\n", "payload\n"),
            1 => ("```\npayload\n\n", "payload\n\n"),
            2 => ("```\n", string.Empty),
            3 => ("```\npayload\n```\n", "payload\n"),
            4 => ("> ```\n> payload\n>\n", "payload\n\n"),
            _ => ("- item\n\n  ```\n  payload\n", "payload\n")
        };
        markdown = markdown.Replace("\n", lineEnding);
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, CommonMark());
        var expected = source.Document.ToHtml(PlainHtml());
        Assert.Contains("<code>" + payload + "</code>", expected);
        AssertReaderModesMatch(markdown, expected);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    public void Lazy_Quote_Breaks_Match_Across_Reader_Modes(int shape) {
        var (markdown, paragraph) = shape switch {
            0 => ("> bar\nbaz\n", "bar\nbaz"),
            1 => ("> > bar\nbaz\n", "bar baz"),
            2 => ("> first\n>\n> bar\nbaz\n", "bar\nbaz"),
            3 => ("> - bar\nbaz\n", "bar baz"),
            4 => ("> bar\nbaz\n> qux\n", "bar\nbaz\nqux"),
            _ => ("> bar  \nbaz\n", "bar<br/>baz")
        };
        var source = MarkdownReader.ParseWithSyntaxTree(markdown, CommonMark());
        var expected = source.Document.ToHtml(PlainHtml());
        Assert.Contains(paragraph, expected);
        AssertReaderModesMatch(markdown, expected);
        MarkdownInvariantAssert.MappedAssociatedObjectsAreConsistent(source);
    }

    private static void AssertReaderModesMatch(string markdown, string expected) {
        Assert.Equal(expected, MarkdownDocument.Parse(markdown, CommonMark()).ToHtml(PlainHtml()));
        Assert.Equal(expected, MarkdownReader.ParseSemantic(markdown, CommonMark()).ToHtml(PlainHtml()));
        Assert.Equal(expected, MarkdownReader.ParseSemanticProjection(markdown, CommonMark()).ToHtml(PlainHtml()));
    }

    private static MarkdownLoadOptions CommonMark() => MarkdownLoadOptions.CreateCommonMarkProfile();

    private static MarkdownHtmlExportOptions PlainHtml() => new() {
        Style = MarkdownHtmlStyle.Plain,
        CssDelivery = MarkdownCssDelivery.None,
        BodyClass = null,
        AutoHeadingIdentifiers = false,
        EscapeNonAsciiText = false,
        NormalizeUrlHostsToIdn = false
    };
}
