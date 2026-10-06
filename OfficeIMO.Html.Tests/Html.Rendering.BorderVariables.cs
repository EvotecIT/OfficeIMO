using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("border:2px solid var(--line);border-color:var(--tone)")]
    [InlineData("border:2px solid red;border-color:var(--tone)")]
    public void HtmlBorders_VariableColorOverrideRetainsTheDeclaredWidthAndStyle(string declarations) {
        string html = "<style>.chip{--line:red;--tone:blue;display:inline-block;" + declarations
            + ";padding:2px;border-radius:8px}</style><p><span class='chip'>Report status</span></p>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlComputedStyle style = Assert.Single(HtmlComputedStyleEngine.Compute(source),
            pair => pair.Key.GetAttribute("class") == "chip").Value;
        Assert.Equal("2px", style.GetValue("border-top-width"));
        Assert.Equal("solid", style.GetValue("border-top-style"));
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(source);
        HtmlRenderShape border = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip" && shape.Shape.StrokeWidth > 0D);
        Assert.Equal(2D, border.Shape.StrokeWidth);
        Assert.Equal(OfficeColor.Blue, border.Shape.StrokeColor);
    }

    [Fact]
    public void HtmlBorders_VariableColorOverrideAcrossRulesKeepsThePrintBorder() {
        const string html = "<style>.chip{--line:red;--tone:blue;border:2px solid var(--line);display:inline-block}"
            + "@media print{.chip[data-state=down]{border-color:var(--tone)}}</style>"
            + "<p><span class='chip' data-state='down'>Down</span></p>";
        HtmlConversionDocument source = HtmlConversionDocument.Parse(html);
        HtmlRenderDocument screen = HtmlRenderEngine.Render(source);
        HtmlRenderDocument print = HtmlRenderEngine.Render(source,new HtmlRenderOptions { Mode=HtmlRenderMode.Paged });
        HtmlRenderShape screenBorder = Assert.Single(screen.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip" && shape.Shape.StrokeWidth > 0D);
        HtmlRenderShape printBorder = Assert.Single(print.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip" && shape.Shape.StrokeWidth > 0D);
        Assert.Equal(OfficeColor.Red, screenBorder.Shape.StrokeColor);
        Assert.Equal(OfficeColor.Blue, printBorder.Shape.StrokeColor);
        Assert.Equal(2D, printBorder.Shape.StrokeWidth);
    }
}
