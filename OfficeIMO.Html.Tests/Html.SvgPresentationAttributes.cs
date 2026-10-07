using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlSvgPresentationAttributes {
    [Fact]
    public void SvgFontAndPaintAttributesOverrideInheritedHtmlStylesAndReachDescendants() {
        const string html = """
            <div style="font-family:Calibri;font-size:16px;font-weight:400;fill:red">
              <svg font-family="Arial" font-size="20" fill="blue">
                <g font-weight="600">
                  <text id="label" font-family="Courier" font-size="10.5" fill="green">Label</text>
                  <text id="inherited">Inherited</text>
                </g>
              </svg>
            </div>
            """;
        var styles = HtmlComputedStyleEngine.Compute(html);
        HtmlComputedStyle label = styles.Single(p => p.Key.Id == "label").Value;
        HtmlComputedStyle inherited = styles.Single(p => p.Key.Id == "inherited").Value;
        Assert.Equal("Courier", label.Properties["font-family"]);
        Assert.Equal("10.5px", label.Properties["font-size"]);
        Assert.Equal("600", label.Properties["font-weight"]);
        Assert.Equal("green", label.Properties["fill"]);
        Assert.Equal(7.875D, label.ResolvedFontSizePoints!.Value, 3);
        Assert.Equal("Arial", inherited.Properties["font-family"]);
        Assert.Equal("20px", inherited.Properties["font-size"]);
        Assert.Equal("600", inherited.Properties["font-weight"]);
        Assert.Equal("blue", inherited.Properties["fill"]);
    }

    [Theory]
    [InlineData("* { font-family:Courier;fill:green }", "", "Courier", "rgba(0, 128, 0, 1)")]
    [InlineData("@layer theme { text { font-family:Courier;fill:green } }", "", "Courier", "rgba(0, 128, 0, 1)")]
    [InlineData("", "font-family:Courier;fill:green", "Courier", "green")]
    [InlineData("text { font-family:Courier!important;fill:green!important }", "font-family:Arial;fill:blue", "Courier", "rgba(0, 128, 0, 1)")]
    [InlineData("text { font-family:inherit;fill:inherit }", "", "Calibri", "red")]
    public void AuthorCssOverridesSvgPresentationAttributes(string css, string inline, string family, string fill) {
        string html = $"<style>{css}</style><div style='font-family:Calibri;fill:red'><svg><text id='label' font-family='Arial' fill='blue' style='{inline}'>Label</text></svg></div>";
        HtmlComputedStyle style = HtmlComputedStyleEngine.Compute(html).Single(p => p.Key.Id == "label").Value;
        Assert.Equal(family, style.Properties["font-family"]);
        Assert.Equal(fill, style.Properties["fill"]);
    }

    [Fact]
    public void HtmlAndNonPresentationAttributesDoNotBecomeCssDeclarations() {
        const string html = """
            <div style="font-family:Calibri;line-height:2">
              <span id="html" font-family="Courier">HTML</span>
              <svg line-height="3"><text id="svg" font-family="Arial !important">SVG</text></svg>
            </div>
            """;
        var styles = HtmlComputedStyleEngine.Compute(html);
        Assert.Equal("Calibri", styles.Single(p => p.Key.Id == "html").Value.Properties["font-family"]);
        HtmlComputedStyle svg = styles.Single(p => p.Key.Id == "svg").Value;
        Assert.Equal("Calibri", svg.Properties["font-family"]);
        Assert.Equal("2", svg.Properties["line-height"]);
    }

    [Fact]
    public void SvgPresentationAttributesRespectTheCssDeclarationBudget() {
        var options = new HtmlConversionDocumentOptions { Trust = HtmlInputTrust.Untrusted };
        options.Limits.MaxCssDeclarations = 1;
        var document = HtmlConversionDocument.Parse("<svg font-family='Arial' font-weight='600'></svg>", options);
        var exception = Assert.Throws<HtmlDomLimitException>(() => HtmlComputedStyleEngine.Compute(document));
        Assert.Equal(HtmlConversionDiagnosticCodes.CssDeclarationLimitExceeded, exception.Code);
    }

}
