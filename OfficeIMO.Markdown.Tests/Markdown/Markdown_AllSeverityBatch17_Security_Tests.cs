using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class MarkdownAllSeverityBatch17SecurityTests {
    [Theory]
    [InlineData("ms-msdt:diagnostic")]
    [InlineData("search-ms:query=secret")]
    [InlineData("shell:Downloads")]
    [InlineData("smb://server/share")]
    public void HtmlRendererRejectsUnlistedLinkSchemes(string target) {
        MarkdownDoc document = MarkdownReader.Parse($"[open]({target})");

        string html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.DoesNotContain("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("href=", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">open<", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("https://example.test/report")]
    [InlineData("mailto:security@example.test")]
    [InlineData("tel:+12025550123")]
    public void HtmlRendererRetainsExplicitlyAllowedLinkSchemes(string target) {
        MarkdownDoc document = MarkdownReader.Parse($"[open]({target})");

        string html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Contains("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(target, html, StringComparison.Ordinal);
    }
}
