using System.Net;
using System.Text.RegularExpressions;
using Markdig;
using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_LiteralBackslash_Review_Tests {
    [Fact]
    public void LiteralTextDoesNotActivateInsertionOrSuperscriptExtensions() {
        const string value = "++added++ and x^2^ with A+B";
        var table = new TableBlock();
        table.Headers.Add("Value");
        table.Rows.Add(new[] { value });
        string markdown = MarkdownDoc.Create().P(value).Add(table).ToMarkdown(
            new MarkdownWriteOptions { TextEscapingMode = MarkdownTextEscapingMode.LiteralBackslash });
        string html = Markdig.Markdown.ToHtml(markdown, Pipeline());
        Assert.DoesNotContain("<ins>", html);
        Assert.DoesNotContain("<sup>", html);
        Assert.Equal(2, Regex.Matches(PlainText(html), Regex.Escape(value)).Count);
        Assert.Contains("A+B", markdown);
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.Inserted = true;
        options.Superscript = true;
        string ownHtml = MarkdownReader.Parse(markdown, options).ToHtmlFragment();
        Assert.DoesNotContain("<ins>", ownHtml);
        Assert.DoesNotContain("<sup>", ownHtml);
        Assert.Equal(2, Regex.Matches(PlainText(ownHtml), Regex.Escape(value)).Count);
    }

    [Fact]
    public void LiteralTextPreservesLongAmpersandRunsAndNonEntitySuffixes() {
        string value = new string('&', 100_000) + " " + "&" + new string('x', 33) + ";";
        string markdown = MarkdownDoc.Create().P(value).ToMarkdown(
            new MarkdownWriteOptions { TextEscapingMode = MarkdownTextEscapingMode.LiteralBackslash });
        Assert.Equal(value, markdown.Trim());
    }

    private static MarkdownPipeline Pipeline() {
        var builder = new MarkdownPipelineBuilder();
        Markdig.MarkdownExtensions.UsePipeTables(builder);
        Markdig.MarkdownExtensions.UseEmphasisExtras(builder);
        return builder.Build();
    }

    private static string PlainText(string html) => WebUtility.HtmlDecode(Regex.Replace(html, "<[^>]+>", string.Empty));
}
