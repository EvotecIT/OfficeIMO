using System;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using OfficeIMO.Html.Providers;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlOptionalNodeTests {
    private static HtmlDocument Parse(string source) => AngleSharpHtmlParser.Instance.Parse(source, new HtmlParseOptions());

    [Fact]
    public void MissingSelectorResultsRetainAccessibilityFallbacks() {
        HtmlElement? missing = Parse("<p>Content</p>").QuerySelector("img");
        Assert.False(HtmlAccessibilitySemantics.HasRole(missing, "heading"));
        Assert.False(HtmlAccessibilitySemantics.HasEpubType(missing, "chapter"));
        Assert.False(HtmlAccessibilitySemantics.TryGetHeadingLevel(missing, out int level));
        Assert.Equal(0, level);
        Assert.Empty(HtmlAccessibilitySemantics.GetAccessibleName(missing));
        Assert.Empty(HtmlAccessibilitySemantics.GetAccessibleName(missing, true));
        Assert.Empty(HtmlAccessibilitySemantics.GetImageAccessibleName(missing));
        Assert.False(HtmlAccessibilitySemantics.IsAriaHidden(missing));
    }

    [Fact]
    public void MissingSelectorResultsRetainImageAndUrlFallbacks() {
        HtmlElement? missing = Parse("<p>Content</p>").QuerySelector("img");
        Assert.Empty(HtmlImageSourceResolver.ResolveImageSource(missing, null, null));
        Assert.Empty(HtmlImageSourceResolver.ResolveImageSourceCandidates(missing, null, null));
        Assert.Empty(HtmlImageSourceResolver.ResolveImageSourceCandidates(missing, null, null, true, 3));
        Assert.Empty(HtmlImageSourceResolver.ResolvePictureSource(missing, null, null));
        Assert.Empty(HtmlImageSourceResolver.ResolvePictureSourceCandidates(missing, null, null));
        Assert.Empty(HtmlImageSourceResolver.ResolvePictureSourceCandidates(missing, null, null, 3));
        Assert.Empty(HtmlImageSourceResolver.ResolveUrlFromSrcSetAttributes(missing, null, null, "srcset"));
        Assert.Empty(HtmlImageSourceResolver.ResolveNormalizedSrcSetAttributes(missing, null, null, "srcset"));
        Assert.Empty(HtmlImageSourceResolver.ResolveUrlAttributes(missing, null, null, "src"));
    }

    [Fact]
    public void InlineCallbacksAcceptMissingNodesAndStillRejectForeignSnapshots() {
        var options = new HtmlToMarkdownOptions();
        HtmlDocument other = Parse("<b>Other</b>");
        bool completed = false;
        options.InlineElementConverters.Add(new HtmlInlineElementConverter("optional", "Optional nodes", context => {
            if (context.Element.LocalName != "span") return null;
            Assert.Empty(context.ConvertNodesToInlineSequence(context.Element.QuerySelector("missing")?.ChildNodes).Nodes);
            var selected = context.ConvertNodesToInlineSequence(new HtmlNode?[] { null, context.Element.QuerySelector("b"), null });
            Assert.NotEmpty(selected.Nodes);
            Assert.Throws<ArgumentException>(() => context.ConvertNodesToInlineSequence(new HtmlNode?[] { null, other.Body }));
            completed = true;
            return selected.Nodes;
        }));
        string markdown = HtmlConversionDocument.Parse("<p><span><b>Kept</b></span></p>").ToMarkdown(options);
        Assert.True(completed);
        Assert.Contains("**Kept**", markdown);
        Assert.DoesNotContain("Other", markdown);
    }

    [Fact]
    public void BlockCallbacksSkipMissingItemsWithoutAcceptingMissingCollections() {
        var options = new HtmlToMarkdownOptions();
        HtmlDocument other = Parse("<p>Other</p>");
        bool completed = false;
        options.ElementBlockConverters.Add(new HtmlElementBlockConverter("optional", "Optional nodes", context => {
            if (context.Element.LocalName != "section") return null;
            HtmlElement p = context.Element.QuerySelector("p")!;
            var blocks = context.ConvertNodesToBlocks(new HtmlNode?[] { null, p, null });
            Assert.NotEmpty(context.ConvertNodesToInlineSequence(new HtmlNode?[] { null, p.FirstChild }).Nodes);
            Assert.Throws<ArgumentNullException>(() => context.ConvertNodesToBlocks(null!));
            Assert.Throws<ArgumentNullException>(() => context.ConvertNodesToInlineSequence(null!));
            Assert.Throws<ArgumentException>(() => context.ConvertNodesToBlocks(new HtmlNode?[] { null, other.Body }));
            completed = true;
            return blocks;
        }));
        string markdown = HtmlConversionDocument.Parse("<section><p>Kept</p></section>").ToMarkdown(options);
        Assert.True(completed);
        Assert.Contains("Kept", markdown);
        Assert.DoesNotContain("Other", markdown);
    }
}
