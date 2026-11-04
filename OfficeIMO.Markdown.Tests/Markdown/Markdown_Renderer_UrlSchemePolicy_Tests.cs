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
        var opts = new MarkdownLoadOptions { RestrictUrlSchemes = true, AllowedUrlSchemes = new[] { "http", "https" } };
        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse("[x](mailto:user@example.com)", opts);
        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });
        Assert.DoesNotContain("mailto:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Reader_Applies_Scheme_Restrictions_To_Windows_Drive_Paths() {
        var opts = new MarkdownLoadOptions { RestrictUrlSchemes = true, AllowedUrlSchemes = new[] { "https" } };
        var doc = OfficeIMO.Markdown.MarkdownDocument.Parse("[x](C:/secret/file.docx)", opts);
        var html = doc.ToHtml(new MarkdownHtmlExportOptions { Style = MarkdownHtmlStyle.Plain, CssDelivery = MarkdownCssDelivery.None, BodyClass = null, Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment });

        Assert.DoesNotContain("C:/secret/file.docx", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }
}

