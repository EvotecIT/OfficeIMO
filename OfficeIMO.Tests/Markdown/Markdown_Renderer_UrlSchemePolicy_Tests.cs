using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_UrlSchemePolicy_Tests {
    [Fact]
    public void MarkdownRenderer_Defaults_To_Allowlist_Schemes() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](ftp://example.com)");
        Assert.DoesNotContain("ftp://", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Can_Restrict_Unknown_Schemes_When_Enabled() {
        var opts = new MarkdownReaderOptions { RestrictUrlSchemes = true, AllowedUrlSchemes = new[] { "http", "https" } };
        var doc = MarkdownReader.Parse("[x](mailto:user@example.com)", opts);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });
        Assert.DoesNotContain("mailto:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }
}

