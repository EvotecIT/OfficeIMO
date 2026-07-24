using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_RawHtmlHandling_Tests {
    [Fact]
    public void GenericAttributes_DoNotBypassSanitizedHtmlAttributePolicy() {
        const string markdown = "Safe {onclick=\"alert(1)\" style=\"display:none\" href=\"javascript:alert(2)\" srcdoc=\"<script>x</script>\" data-safe=\"yes\" title=\"kept\"}\n\n![safe](/safe.png){srcset=\"https://attacker.test/leak.png 2x\" imagesrcset=\"https://attacker.test/preload.png 1x\"}";
        var readerOptions = new MarkdownReaderOptions { GenericAttributes = true };
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Sanitize
        };

        string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, readerOptions)
            .ToHtmlFragment(htmlOptions);

        Assert.DoesNotContain("onclick", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("style=", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("srcdoc", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("srcset", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("imagesrcset", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-safe=\"yes\"", html, StringComparison.Ordinal);
        Assert.Contains("title=\"kept\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void GenericAttributes_PreserveTrustedAllowRendering() {
        const string markdown = "Trusted {style=\"color:red\" href=\"https://example.test\" onclick=\"trusted()\"}";
        var readerOptions = new MarkdownReaderOptions { GenericAttributes = true };
        var htmlOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Allow
        };

        string html = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, readerOptions)
            .ToHtmlFragment(htmlOptions);

        Assert.Contains("style=\"color:red\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.test\"", html, StringComparison.Ordinal);
        Assert.Contains("onclick=\"trusted()\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Strip_RawHtml_Blocks() {
        var md = "<div>hi</div>\n\nParagraph";
        var opts = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Strip };
        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.DoesNotContain("<div>hi</div>", html, StringComparison.Ordinal);
        Assert.Contains("Paragraph", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Escape_RawHtml_Blocks() {
        var md = "<script>alert(1)</script>";
        var opts = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Escape };
        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlOptions_Can_Escape_RawHtml_With_NonAscii_Text_Literal_For_Markdig_Style_Output() {
        const string md = "Before <span>åinline</span> after\n\n<div>åblock</div>\n\n<!-- åcomment -->";
        var opts = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Escape,
            EscapeNonAsciiText = false
        };

        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.Contains("Before &lt;span&gt;åinline&lt;/span&gt; after", html, StringComparison.Ordinal);
        Assert.Contains("&lt;div&gt;åblock&lt;/div&gt;", html, StringComparison.Ordinal);
        Assert.Contains("&lt;!-- åcomment --&gt;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Sanitize_RawHtml_Blocks_With_Allowlist() {
        var md = "<details open onclick=\"alert(1)\"><summary>Title</summary><script>alert(1)</script><u>ok</u><br></details>";
        var opts = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Sanitize };
        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.Contains("<details open>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<summary>Title</summary>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<u>ok</u>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<br/>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("onclick", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlOptions_Can_Sanitize_RawHtml_With_NonAscii_Text_Literal_For_Markdig_Style_Output() {
        const string md = "<details><summary>åTitle</summary><script>åbad</script><u>åok</u></details>";
        var opts = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Sanitize,
            EscapeNonAsciiText = false
        };

        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.Contains("<summary>åTitle</summary>", html, StringComparison.Ordinal);
        Assert.Contains("<u>åok</u>", html, StringComparison.Ordinal);
        Assert.Contains("&lt;script&gt;åbad&lt;/script&gt;", html, StringComparison.Ordinal);
        Assert.DoesNotContain("&#229;", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Sanitize_Inline_RawHtml_With_NonAscii_Text_Literal() {
        const string md = "Before <span>åbad</span> and <u>åok</u>";
        var opts = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Sanitize,
            EscapeNonAsciiText = false
        };

        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.Contains("Before &lt;span&gt;åbad&lt;/span&gt; and <u>åok</u>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("&#229;", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("type 1 script", "<script>\nalert(1)\n</script>", "<script>", "&lt;script&gt;")]
    [InlineData("type 2 comment", "<!-- keep -->", "<!-- keep -->", "&lt;!-- keep --&gt;")]
    [InlineData("type 3 processing instruction", "<?xml version=\"1.0\"?>", "<?xml", "&lt;?xml")]
    [InlineData("type 4 declaration", "<!DOCTYPE html>", "<!DOCTYPE", "&lt;!DOCTYPE")]
    [InlineData("type 5 CDATA", "<![CDATA[<p>literal</p>]]>", "<![CDATA", "&lt;![CDATA")]
    [InlineData("type 6 block tag", "<div onclick=\"alert(1)\">ok</div>", "<div", "&lt;div")]
    [InlineData("type 7 custom tag", "<custom onclick=\"alert(1)\">\nok</custom>", "<custom", "&lt;custom")]
    public void RawHtmlHandling_Security_Profiles_Cover_CommonMark_Html_Block_Shapes(
        string _,
        string rawMarkdown,
        string unsafeFragment,
        string escapedFragment) {
        string markdown = rawMarkdown + "\n\nParagraph";
        var doc = OfficeIMO.Markdown.MarkdownReader.Parse(markdown);

        var stripped = doc.ToHtmlFragment(CreatePlainHtmlOptions(RawHtmlHandling.Strip));
        Assert.DoesNotContain(unsafeFragment, stripped, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(escapedFragment, stripped, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p>Paragraph</p>", stripped, StringComparison.Ordinal);

        var escaped = doc.ToHtmlFragment(CreatePlainHtmlOptions(RawHtmlHandling.Escape));
        Assert.DoesNotContain(unsafeFragment, escaped, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<pre class=\"md-raw-html\"><code>", escaped, StringComparison.Ordinal);
        Assert.Contains(escapedFragment, escaped, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p>Paragraph</p>", escaped, StringComparison.Ordinal);

        var sanitized = doc.ToHtmlFragment(CreatePlainHtmlOptions(RawHtmlHandling.Sanitize));
        Assert.DoesNotContain(unsafeFragment, sanitized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(escapedFragment, sanitized, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p>Paragraph</p>", sanitized, StringComparison.Ordinal);
    }

    [Fact]
    public void GitHubHtmlTagFilter_Filters_Dangerous_RawHtml_Blocks_And_Inlines_When_RawHtml_Is_Allowed() {
        const string md = "Inline <xmp>bad</xmp> but <strong>ok</strong>.\n\n<script>alert(1)</script>\n\n<custom>ok</custom>";
        var opts = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = RawHtmlHandling.Allow,
            GitHubHtmlTagFilter = true
        };

        var html = OfficeIMO.Markdown.MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.Contains("&lt;xmp>bad&lt;/xmp>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<strong>ok</strong>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script>alert(1)&lt;/script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<custom>ok</custom>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<xmp>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GitHubFlavoredMarkdown_Html_Profile_Enables_Tag_Filter_Without_Security_Stripping() {
        const string markdown = """
- [x] done

Inline <xmp>bad</xmp>.

<script>alert(1)</script>
""";

        var document = OfficeIMO.Markdown.MarkdownReader.Parse(markdown, MarkdownReaderOptions.CreateGitHubFlavoredMarkdownProfile());
        var options = HtmlOptions.CreateGitHubFlavoredMarkdownProfile();

        var html = document.ToHtmlFragment(options);

        Assert.True(options.GitHubTaskListHtml);
        Assert.True(options.GitHubFootnoteHtml);
        Assert.True(options.GitHubHtmlTagFilter);
        Assert.Equal(RawHtmlHandling.Allow, options.RawHtmlHandling);
        Assert.Contains("<input type=\"checkbox\" checked=\"\" disabled=\"\" /> done", html, StringComparison.Ordinal);
        Assert.Contains("&lt;xmp>bad&lt;/xmp>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script>alert(1)&lt;/script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void GitHubFlavoredMarkdown_Html_Profile_Does_Not_Change_Strict_Renderer_Security_Defaults() {
        var strict = MarkdownRendererPresets.CreateStrict(MarkdownReaderOptions.MarkdownDialectProfile.GitHubFlavoredMarkdown);

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(
            "<script>alert(1)</script>\n\n- [x] done",
            strict);

        Assert.Equal(RawHtmlHandling.Strip, strict.HtmlOptions.RawHtmlHandling);
        Assert.False(strict.HtmlOptions.GitHubHtmlTagFilter);
        Assert.DoesNotContain("script", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("done", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_To_Stripping_RawHtml() {
        var md = "<div>hi</div>";
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, new MarkdownRendererOptions {
            ReaderOptions = new MarkdownReaderOptions { HtmlBlocks = true, InlineHtml = true }
        });

        Assert.DoesNotContain("<div>hi</div>", html, StringComparison.Ordinal);
    }

    private static HtmlOptions CreatePlainHtmlOptions(RawHtmlHandling handling) =>
        new() {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null,
            RawHtmlHandling = handling
        };
}
