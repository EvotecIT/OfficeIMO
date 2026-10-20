using AngleSharp.Html.Parser;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPageMarginParserLimits_TrustedInputDoesNotAcquireAnIndependentEightMegabyteCap() {
        string declarations = "/*" + new string(' ', 8 * 1024 * 1024) + "*/font:16px/10cqw Pinned";
        using var document = new HtmlParser().ParseDocument(PageMarginTypographySource(declarations));
        var options = PageMarginTypographyOptions();

        HtmlCssPageRuleSet rules = HtmlCssPageSettingsResolver.Apply(document, options, new HtmlDiagnosticReport());

        Assert.Equal("10cqw", Assert.Single(rules.ResolveMarginBoxes(
            1, null, rules.ResolveGeometry(1, null, options), options)).Value.LineHeight);
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
