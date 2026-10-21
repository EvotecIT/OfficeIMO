using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPageMarginParserLimits_TrustedInputDoesNotAcquireAnIndependentEightMegabyteCap() {
        string declarations = "/*" + new string(' ', 8 * 1024 * 1024) + "*/font:16px/10cqw Pinned";
        HtmlConversionDocument document = HtmlConversionDocument.Parse(PageMarginTypographySource(declarations),
            HtmlConversionDocumentOptions.CreateTrustedProfile());
        var options = PageMarginTypographyOptions();
        options.PageSize = new OfficeIMO.Drawing.OfficePageSize(360D / 96D, 480D / 96D);

        HtmlRenderPage page = Assert.Single(HtmlRenderEngine.Render(document, options).Pages);

        Assert.Equal(36D, PageMarginTypographyText(page).LineHeight, 6);
    }

    [Fact]
    public void HtmlPageMarginParserLimits_ConfiguredTokenLimitUsesTheSharedDiagnostic() {
        using var document = new HtmlParser().ParseDocument(PageMarginTypographySource("font:16px/10cqw Pinned"));
        var limits = HtmlConversionLimits.CreateTrustedProfile();
        limits.MaxCssTokens = 4;

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlCssPageSettingsResolver.Apply(document, PageMarginTypographyOptions(), new HtmlDiagnosticReport(), limits));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssTokenLimitExceeded, exception.Code);
    }
}
