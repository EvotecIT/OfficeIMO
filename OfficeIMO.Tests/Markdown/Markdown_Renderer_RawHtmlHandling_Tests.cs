using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_RawHtmlHandling_Tests {
    [Fact]
    public void HtmlOptions_Can_Strip_RawHtml_Blocks() {
        var md = "<div>hi</div>\n\nParagraph";
        var opts = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Strip };
        var html = MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.DoesNotContain("<div>hi</div>", html, StringComparison.Ordinal);
        Assert.Contains("Paragraph", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Escape_RawHtml_Blocks() {
        var md = "<script>alert(1)</script>";
        var opts = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null, RawHtmlHandling = RawHtmlHandling.Escape };
        var html = MarkdownReader.Parse(md).ToHtmlFragment(opts);

        Assert.DoesNotContain("<script>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script&gt;alert(1)&lt;/script&gt;", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_To_Stripping_RawHtml() {
        var md = "<div>hi</div>";
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, new MarkdownRendererOptions {
            ReaderOptions = new MarkdownReaderOptions { HtmlBlocks = true, InlineHtml = true }
        });

        Assert.DoesNotContain("<div>hi</div>", html, StringComparison.Ordinal);
    }
}

