using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRender_ChargesSelectorsOnlyForTheirElementOrPseudoTarget() {
        const string html = """
            <style>
              #target { color:#123456; }
              #target::before { content:"Before "; }
              #target:after { content:" After"; }
            </style>
            <p id="target">Body</p>
            """;
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 3;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html,
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(source);
        HtmlRenderText[] text = Assert.Single(rendered.Pages).Visuals.OfType<HtmlRenderText>().ToArray();

        Assert.Equal(new[] { "Before ", "Body", " After" }, text.Select(item => item.Text));
        Assert.All(text, item => Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), item.Color));
        Assert.Equal("p#target::before", text[0].Source);
        Assert.Equal("p#target::after", text[2].Source);
        Assert.DoesNotContain(rendered.Diagnostics,
            item => item.Code == HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded);
    }

    [Theory]
    [InlineData("")]
    [InlineData("::before")]
    public void HtmlRender_StillBoundsUnsuccessfulCandidateMatches(string pseudo) {
        string html = "<style>#target.unmatched" + pseudo + "{color:red;content:'wrong'}"
            + "#target" + pseudo + "{color:blue;content:'selected'}</style><p id='target'>Body</p>";
        HtmlConversionLimits limits = HtmlConversionLimits.CreateUntrustedProfile();
        limits.MaxSelectorEvaluations = 1;
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html,
            new HtmlConversionDocumentOptions { Limits = limits });

        HtmlDomLimitException error = Assert.Throws<HtmlDomLimitException>(
            () => HtmlRenderTestDriver.Render(source));

        Assert.Equal(HtmlConversionDiagnosticCodes.CssSelectorEvaluationLimitExceeded, error.Code);
        Assert.Equal(nameof(HtmlConversionLimits.MaxSelectorEvaluations), error.LimitSource);
    }
}
