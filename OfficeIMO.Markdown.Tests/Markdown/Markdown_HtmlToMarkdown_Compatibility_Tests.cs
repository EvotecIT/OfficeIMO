using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class MarkdownHtmlToMarkdownCompatibilityTests {
    [Fact]
    public void HtmlToMarkdown_SmartHref_EmitsPlainTextForSelfLinks() {
        const string html = """
<p>
  <a href="https://example.com">https://example.com</a>
  <a href="mailto:user@example.com">user@example.com</a>
  <a href="tel:+15551234567">+15551234567</a>
  <a href="https://example.com/docs">Docs</a>
</p>
""";

        string markdown = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            SmartHref = true
        }));

        Assert.Contains("https://example.com user@example.com [+15551234567](tel:+15551234567) [Docs](https://example.com/docs)", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[https://example.com]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[user@example.com]", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_FiltersElementsWithSelectorsAndPredicates() {
        const string html = """
<article>
  <p>Keep</p>
  <aside class="ad">Advertisement</aside>
  <p data-skip="true">Drop me</p>
</article>
""";

        var options = new HtmlToMarkdownOptions();
        options.ExcludeSelectors.Add(".ad");
        options.ElementFilters.Add(element => string.Equals(element.GetAttribute("data-skip"), "true", StringComparison.Ordinal));

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("Keep", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_FiltersBaseElementBeforeResolvingRelativeUrls() {
        const string html = """
<html>
  <head><base href="https://external.example/"></head>
  <body><p><a href="guide">Guide</a></p></body>
</html>
""";

        var options = new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://docs.example/root/")
        };
        options.ExcludeSelectors.Add("base");

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("[Guide](https://docs.example/root/guide)", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_TagAliases_MapUnsupportedTagsToBuiltInConverters() {
        const string html = "<p><highlight>Important</highlight> and <custom-bold>bold</custom-bold></p>";

        var options = new HtmlToMarkdownOptions {
            PreserveUnsupportedInlineHtml = false
        };
        options.TagAliases["highlight"] = "mark";
        options.TagAliases["custom-bold"] = "strong";

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("==Important== and **bold**", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_TagAliases_MapStructuredChildrenToBuiltInConverters() {
        const string html = "<custom-list><custom-item>First</custom-item><custom-item><strong>Second</strong></custom-item></custom-list>";

        var options = new HtmlToMarkdownOptions();
        options.TagAliases["custom-list"] = "ul";
        options.TagAliases["custom-item"] = "li";

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("- First\n- **Second**", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_TagAliases_MapMediaChildrenToBuiltInConverters() {
        const string html = """
<custom-figure>
  <custom-picture>
    <custom-img src="media/photo.png" alt="Photo">
  </custom-picture>
  <custom-caption>Caption text</custom-caption>
</custom-figure>
""";

        var options = new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        };
        options.TagAliases["custom-figure"] = "figure";
        options.TagAliases["custom-picture"] = "picture";
        options.TagAliases["custom-img"] = "img";
        options.TagAliases["custom-caption"] = "figcaption";

        var document = new HtmlToMarkdownConverter().ConvertToDocument(html, options);
        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));

        Assert.Equal("https://example.com/docs/media/photo.png", image.Path);
        Assert.Equal("Photo", image.Alt);
        Assert.Equal("Caption text", image.Caption);
    }

    [Fact]
    public void HtmlToMarkdown_TagAliases_MapPictureSourcesToPreferredImagePath() {
        const string html = """
<custom-picture>
  <custom-source srcset="large.png 2x">
  <custom-img src="small.png" alt="Photo">
</custom-picture>
""";

        var options = new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        };
        options.TagAliases["custom-picture"] = "picture";
        options.TagAliases["custom-source"] = "source";
        options.TagAliases["custom-img"] = "img";

        var document = new HtmlToMarkdownConverter().ConvertToDocument(html, options);
        var image = Assert.IsType<ImageBlock>(Assert.Single(document.Blocks));
        var source = Assert.Single(image.PictureSources);

        Assert.Equal("https://example.com/docs/large.png", image.Path);
        Assert.Equal("https://example.com/docs/small.png", image.PictureFallbackPath);
        Assert.Equal("https://example.com/docs/large.png", source.Path);
        Assert.Equal("Photo", image.Alt);
    }

    [Fact]
    public void HtmlToMarkdown_PassThroughTags_PreserveOriginalHtmlEvenForKnownTags() {
        const string html = "<p>Keep <strong>literal</strong></p>";

        var options = new HtmlToMarkdownOptions();
        options.PassThroughTags.Add("strong");

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("Keep <strong>literal</strong>", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_UnknownBlockHandling_CanBypassDropPreserveOrRaise() {
        const string html = "<x-card><p>Inner <strong>text</strong></p></x-card>";

        string bypassed = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownBlockHandling = HtmlUnknownTagHandling.Bypass
        }));
        string dropped = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownBlockHandling = HtmlUnknownTagHandling.Drop
        }));
        string preserved = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownBlockHandling = HtmlUnknownTagHandling.Preserve
        }));
        string bypassedInlineOnly = Normalize("<custom-widget><strong>Custom</strong> payload</custom-widget>".ToMarkdown(new HtmlToMarkdownOptions {
            UnknownBlockHandling = HtmlUnknownTagHandling.Bypass
        }));
        string droppedInlineOnly = Normalize("<custom-widget><strong>Custom</strong> payload</custom-widget>".ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = false,
            UnknownBlockHandling = HtmlUnknownTagHandling.Drop
        }));

        Assert.Equal("Inner **text**", bypassed);
        Assert.Equal("**Custom** payload", bypassedInlineOnly);
        Assert.Equal(string.Empty, droppedInlineOnly);
        Assert.Equal(string.Empty, dropped);
        Assert.Contains("<x-card>", preserved, StringComparison.Ordinal);
        Assert.Throws<NotSupportedException>(() => html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownBlockHandling = HtmlUnknownTagHandling.Raise
        }));
        Assert.Throws<NotSupportedException>(() => "<custom-widget>payload</custom-widget>".ToMarkdown(new HtmlToMarkdownOptions {
            PreserveUnsupportedBlocks = false,
            UnknownBlockHandling = HtmlUnknownTagHandling.Raise
        }));
    }

    [Fact]
    public void HtmlToMarkdown_UnknownInlineHandling_CanBypassDropPreserveOrRaise() {
        const string html = "<p>Before <x-chip>inside</x-chip> after</p>";

        string bypassed = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownInlineHandling = HtmlUnknownTagHandling.Bypass
        }));
        string dropped = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownInlineHandling = HtmlUnknownTagHandling.Drop
        }));
        string preserved = Normalize(html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownInlineHandling = HtmlUnknownTagHandling.Preserve
        }));

        Assert.Equal("Before inside after", bypassed);
        Assert.Contains("Before", dropped, StringComparison.Ordinal);
        Assert.Contains("after", dropped, StringComparison.Ordinal);
        Assert.DoesNotContain("inside", dropped, StringComparison.Ordinal);
        Assert.DoesNotContain("x-chip", dropped, StringComparison.Ordinal);
        Assert.Equal("Before <x-chip>inside</x-chip> after", preserved);
        Assert.Throws<NotSupportedException>(() => html.ToMarkdown(new HtmlToMarkdownOptions {
            UnknownInlineHandling = HtmlUnknownTagHandling.Raise
        }));
    }

    [Fact]
    public void HtmlToMarkdown_GitHubProfile_UsesSmartHrefAndPipeTables() {
        const string html = """
<article>
  <p><a href="https://example.com">https://example.com</a></p>
  <table><tr><th>Name</th><th>Value</th></tr><tr><td>Area</td><td>Markdown</td></tr></table>
</article>
""";

        string markdown = Normalize(html.ToMarkdown(HtmlToMarkdownOptions.CreateGitHubFlavoredMarkdownProfile()));

        Assert.Contains("https://example.com", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("[https://example.com]", markdown, StringComparison.Ordinal);
        Assert.Contains("| Name | Value |", markdown, StringComparison.Ordinal);
        Assert.Contains("| --- | --- |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_CommonMarkProfile_UsesRawHtmlForTablesAndKeepsExplicitSelfLinks() {
        const string html = """
<article>
  <p><a href="https://example.com">https://example.com</a></p>
  <p><del>old</del> and <mark>new</mark></p>
  <ul><li><input type="checkbox" checked>Done</li><li><input type="checkbox">Open</li></ul>
  <table><tr><th>Name</th><th>Value</th></tr><tr><td>Area</td><td>Markdown</td></tr></table>
</article>
""";

        string markdown = Normalize(html.ToMarkdown(HtmlToMarkdownOptions.CreateCommonMarkProfile()));

        Assert.Contains("[https://example.com](https://example.com)", markdown, StringComparison.Ordinal);
        Assert.Contains("<del>old</del>", markdown, StringComparison.Ordinal);
        Assert.Contains("<mark>new</mark>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("~~old~~", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("==new==", markdown, StringComparison.Ordinal);
        Assert.Contains("<input", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- [x]", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("- [ ]", markdown, StringComparison.Ordinal);
        Assert.Contains("<table>", markdown, StringComparison.Ordinal);
        Assert.Contains("<th>Name</th>", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("| Name | Value |", markdown, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlToMarkdown_CommonMarkProfile_EscapesLineStartsThatWouldBecomeBlocks() {
        const string html = "<p># not heading</p><p>- not list</p>";

        string markdown = Normalize(html.ToMarkdown(HtmlToMarkdownOptions.CreateCommonMarkProfile()));

        Assert.Equal("\\# not heading\n\n\\- not list", markdown);
    }

    [Fact]
    public void HtmlToMarkdown_TagAliases_UnwrapRejectedAnchorHrefs() {
        const string html = "<custom-link href=\"javascript:alert(1)\"><img src=\"ok.png\" alt=\"Ok\"></custom-link>";

        var options = new HtmlToMarkdownOptions {
            BaseUri = new Uri("https://example.com/docs/")
        };
        options.TagAliases["custom-link"] = "a";

        string markdown = Normalize(html.ToMarkdown(options));

        Assert.Equal("![Ok](https://example.com/docs/ok.png)", markdown);
        Assert.DoesNotContain("javascript:", markdown, StringComparison.OrdinalIgnoreCase);
    }

    private static string Normalize(string value) {
        return value.Replace("\r\n", "\n").Trim();
    }
}
