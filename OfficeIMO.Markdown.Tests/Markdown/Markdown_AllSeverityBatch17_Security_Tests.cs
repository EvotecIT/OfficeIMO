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

    [Fact]
    public void HtmlRendererAllowsAdditionalSchemesOnlyWhenExplicitlyConfigured() {
        MarkdownDoc document = MarkdownReader.Parse("[open](acme-safe:resource)");
        var defaultOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };

        string defaultHtml = document.ToHtmlFragment(defaultOptions);
        Assert.DoesNotContain("href=", defaultHtml, StringComparison.OrdinalIgnoreCase);

        var configuredOptions = new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
        configuredOptions.AdditionalAllowedLinkSchemes.Add("acme-safe");

        string configuredHtml = document.ToHtmlFragment(configuredOptions);
        Assert.Contains("href=\"acme-safe:resource\"", configuredHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendererKeepsColonContainingRelativeTargetsThatCannotNameAProtocol() {
        MarkdownDoc document = MarkdownReader.Parse("[link](foo\\)\\:)");

        string html = document.ToHtmlFragment(new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        });

        Assert.Contains("href=\"foo):\"", html, StringComparison.Ordinal);
    }
}
