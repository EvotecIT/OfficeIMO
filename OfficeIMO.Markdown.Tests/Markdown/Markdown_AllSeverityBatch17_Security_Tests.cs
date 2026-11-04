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
        MarkdownDocument document = MarkdownDocument.Parse($"[open]({target})");

        string html = document.ToHtml(new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null,

            Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment
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
        MarkdownDocument document = MarkdownDocument.Parse($"[open]({target})");

        string html = document.ToHtml(new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null,

            Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment
        });

        Assert.Contains("<a ", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(target, html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendererAllowsAdditionalSchemesOnlyWhenExplicitlyConfigured() {
        MarkdownDocument document = MarkdownDocument.Parse("[open](acme-safe:resource)");
        var defaultOptions = new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null
        };

        string defaultHtml = document.ToHtml(defaultOptions);
        Assert.DoesNotContain("href=", defaultHtml, StringComparison.OrdinalIgnoreCase);

        var configuredOptions = new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null
        };
        configuredOptions.AdditionalAllowedLinkSchemes.Add("acme-safe");

        string configuredHtml = document.ToHtml(configuredOptions);
        Assert.Contains("href=\"acme-safe:resource\"", configuredHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRendererKeepsColonContainingRelativeTargetsThatCannotNameAProtocol() {
        MarkdownDocument document = MarkdownDocument.Parse("[link](foo\\)\\:)");

        string html = document.ToHtml(new MarkdownHtmlExportOptions
        {
            Style = MarkdownHtmlStyle.Plain,
            CssDelivery = MarkdownCssDelivery.None,
            BodyClass = null,

            Kind = global::OfficeIMO.Markdown.MarkdownHtmlKind.Fragment
        });

        Assert.Contains("href=\"foo):\"", html, StringComparison.Ordinal);
    }
}
