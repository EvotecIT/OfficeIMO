using System.Net;
using System.Text.RegularExpressions;
using Markdig;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_LiteralBackslash_Writer_Tests {
    private static MarkdownSaveOptions LiteralOptions() => new() { TextEscapingMode = MarkdownTextEscapingMode.LiteralBackslash };

    [Fact]
    public void ModeIsOptInClonesAndDoesNotLeakIntoOtherWrites() {
        var document = MarkdownDocument.Create().P("A (B) & C");
        Assert.Equal(@"A \(B\) & C", document.ToMarkdown().Trim());
        var options = LiteralOptions().Clone();
        Assert.Equal(MarkdownTextEscapingMode.LiteralBackslash, options.TextEscapingMode);
        Assert.Equal("A (B) & C", document.ToMarkdown(options).Trim());
        Assert.Equal(@"A \(B\) & C", document.ToMarkdown().Trim());
    }

    [Fact]
    public void NativeTextHeadingAndEmphasisUseBackslashesWithoutEntitiesOrDoubleEscaping() {
        const string value = "Read (owner) <img src=x> &amp; &#60; &#x3c; A & B [link](https://example.test) *text* _value_ `code` | \\ $5 ^2!";
        const string escaped = @"Read (owner) \<img src=x\> \&amp; \&\#60; \&\#x3c; A & B \[link\](https://example.test) \*text\* \_value\_ \`code\` \| \\ $5 ^2!";
        var document = MarkdownDocument.Create().H1(value).P(value).P(paragraph => paragraph.Bold(value));
        string markdown = document.ToMarkdown(LiteralOptions()).Replace("\r\n", "\n");
        Assert.Equal("# " + escaped + "\n\n" + escaped + "\n\n**" + escaped + "**\n", markdown);
        string html = Markdig.Markdown.ToHtml(markdown);
        Assert.DoesNotContain("<img ", html);
        Assert.DoesNotContain("<a href=", html);
        Assert.Contains(value, PlainText(html));
    }

    [Fact]
    public void NativeTableCellsRetainPipesBackslashesParenthesesAndLiteralHtml() {
        const string value = "A | B \\ (owner) <b>text</b> &lt; [link](target) *bold*";
        const string escaped = @"A \| B \\ (owner) \<b\>text\</b\> \&lt; \[link\](target) \*bold\*";
        var table = new TableBlock();
        table.Headers.AddRange(new[] { "Value", "Other" });
        table.Rows.Add(new[] { value, "Kept" });
        table.Rows.Add(new[] { "Short row" });
        var document = MarkdownDocument.Create().Add(table);
        string markdown = document.ToMarkdown(LiteralOptions());
        Assert.Contains("| " + escaped + " | Kept |", markdown);
        Assert.Contains("| Short row |  |", markdown);
        var builder = new MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        string html = Markdig.Markdown.ToHtml(markdown, builder.Build());
        Assert.Equal(2, Regex.Matches(html, "<th>").Count);
        Assert.Equal(4, Regex.Matches(html, "<td>").Count);
        Assert.DoesNotContain("<b>", html);
        Assert.DoesNotContain("<a href=", html);
        Assert.Contains(value, PlainText(html));
    }

    [Fact]
    public void ReferenceLikeLabelsKeepLiteralColonsInParagraphsListsAndTableCells() {
        const string value = "[Policy]: none";
        var table = new TableBlock();
        table.Headers.Add("Value");
        table.Rows.Add(new[] { value });
        var document = MarkdownDocument.Create().P(value).Ul(list => list.Item(value)).Add(table);
        string markdown = document.ToMarkdown(LiteralOptions());
        Assert.DoesNotContain("&#58;", markdown);
        Assert.Equal(3, Regex.Matches(markdown, Regex.Escape(@"\[Policy\]: none")).Count);
        Assert.DoesNotContain(@"\:", markdown);
        var pipeline = new MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(pipeline);
        string html = Markdig.Markdown.ToHtml(markdown, pipeline.Build());
        Assert.DoesNotContain("<a ", html);
        Assert.Equal(3, Regex.Matches(PlainText(html), Regex.Escape(value)).Count);
    }

    [Fact]
    public void ParsedTableRetainsExplicitMarkdownFormatting() {
        const string source = "| Value |\n| --- |\n| **bold** and (owner) |\n";
        var document = MarkdownDocument.Parse(source);
        string markdown = document.ToMarkdown(LiteralOptions());
        Assert.Contains("| **bold** and (owner) |", markdown);
        Assert.Contains("<strong>bold</strong>", MarkdownDocument.Parse(markdown).ToHtml());
    }

    [Fact]
    public void DecodedAndLineStartTextNodesHonorTheLiteralMode() {
        var decoded = new InlineSequence();
        decoded.ReplaceItems(new IMarkdownInline[] { new DecodedHtmlEntityTextRun("<u>(owner)</u> &amp;") });
        var escaped = new InlineSequence();
        escaped.ReplaceItems(new IMarkdownInline[] { new LineStartEscapedTextRun("1. (owner)") });
        var document = MarkdownDocument.Create().Add(new ParagraphBlock(decoded)).Add(new ParagraphBlock(escaped));
        string markdown = document.ToMarkdown(LiteralOptions()).Replace("\r\n", "\n");
        Assert.Equal(@"\<u\>(owner)\</u\> \&amp;" + "\n\n" + @"1\. (owner)" + "\n", markdown);
    }

    [Theory]
    [InlineData("# Heading", @"\# Heading")]
    [InlineData("- List", @"\- List")]
    [InlineData("+ List", @"\+ List")]
    [InlineData("1. List", @"1\. List")]
    [InlineData("1) List", @"1\) List")]
    public void LiteralParagraphTextCannotIntroduceHeadingsOrLists(string value, string expected) {
        string markdown = MarkdownDocument.Create().P(value).ToMarkdown(LiteralOptions());
        Assert.Equal(expected, markdown.Trim());
        Assert.Equal("<p>" + value + "</p>", Markdig.Markdown.ToHtml(markdown).Trim());
    }

    private static string PlainText(string html) => WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", string.Empty));
}
