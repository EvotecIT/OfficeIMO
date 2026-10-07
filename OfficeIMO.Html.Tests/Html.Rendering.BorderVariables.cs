using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("border-left-width:calc(1em - 20px)", "left", false)]
    [InlineData("border-left-width:calc(1em - 20px)", "left", true)]
    [InlineData("border-left:calc(1em - 20px) solid green", "left", false)]
    [InlineData("border-width:calc(1em - 20px)", "", false)]
    [InlineData("border:calc(1em - 20px) solid green", "", false)]
    public void HtmlBorders_LengthMathUsesTheElementContextAfterSyntaxValidation(string declaration, string side, bool inline) {
        string declarations = "display:inline-block;font-size:32px;border:2px solid green;" + declaration;
        string html = inline
            ? "<p><span class='chip' style='" + declarations + "'>Status</span></p>"
            : "<style>.chip{" + declarations + "}</style><p><span class='chip'>Status</span></p>";
        HtmlRenderShape border = Assert.Single(HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == (side.Length == 0 ? "span.chip" : "span.chip:border-" + side) && shape.Shape.StrokeWidth > 0D);
        Assert.Equal(12D, border.Shape.StrokeWidth);
        Assert.Equal(OfficeColor.Green, border.Shape.StrokeColor);
    }

    [Theory]
    [InlineData("border-color:'red'")]
    [InlineData("border-color:'red'!important")]
    [InlineData("border-style:'solid'")]
    [InlineData("border-width:'6px'")]
    [InlineData("border-left-color:'red'")]
    [InlineData("border-left-style:'solid'")]
    [InlineData("border-left-width:'6px'")]
    [InlineData("border-color:invert")]
    [InlineData("border-width:6")]
    [InlineData("border-width:25%")]
    [InlineData("border-width:calc(25% + 1px)")]
    public void HtmlBorders_InvalidRawComponentTokensDoNotReplaceTheValidBorder(string declaration) {
        string html = "<style>.chip{display:inline-block;border:2px solid green;" + declaration
            + "}</style><p><span class='chip'>Status</span></p>";
        HtmlRenderShape border = Assert.Single(HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip" && shape.Shape.StrokeWidth > 0D);
        Assert.Equal(2D, border.Shape.StrokeWidth);
        Assert.Equal(OfficeColor.Green, border.Shape.StrokeColor);
    }

    [Fact]
    public void HtmlBorders_SupportsUsesTheSameRawSyntaxAsOrdinaryDeclarations() {
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border-left-width:calc(1em - 20px))"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(border:calc(1em - 20px) solid green)"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-color:'red')"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-style:'solid')"));
        Assert.False(HtmlComputedStyleEngine.IsApplicableSupports("(border-width:'6px')"));
    }

    [Theory]
    [InlineData("top", false)]
    [InlineData("right", false)]
    [InlineData("bottom", false)]
    [InlineData("left", false)]
    [InlineData("top", true)]
    [InlineData("right", true)]
    [InlineData("bottom", true)]
    [InlineData("left", true)]
    public void HtmlBorders_PhysicalSideShorthandsRespectLaterColorAndWidthDeclarations(string side, bool important) {
        string declarations = important
            ? $"border:2px solid green;border-{side}:4px solid red;border-{side}-color:var(--tone)!important;border:6px solid green"
            : $"border:2px solid green;border-{side}:2px solid red;border-{side}-color:var(--tone);border-{side}:4px solid green";
        string html = "<style>.chip{display:inline-block;--tone:blue;" + declarations
            + "}</style><p><span class='chip'>Status</span></p>";
        HtmlRenderShape edge = Assert.Single(HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip:border-" + side);
        Assert.Equal(important ? 6D : 4D, edge.Shape.StrokeWidth);
        Assert.Equal(important ? OfficeColor.Blue : OfficeColor.Green, edge.Shape.StrokeColor);
    }

    [Theory]
    [InlineData("print")]
    [InlineData("layer")]
    [InlineData("pseudo")]
    public void HtmlBorders_AuthoredDeclarationOrderReachesPrintLayersAndGeneratedBoxes(string route) {
        const string declarations = "display:inline-block;--tone:blue;border:2px solid red;border-color:var(--tone);border:4px solid green";
        string rule = route == "pseudo" ? ".chip::before{content:'Status';" + declarations + "}" : ".chip{" + declarations + "}";
        if (route == "print") rule = "@media print{" + rule + "}";
        if (route == "layer") rule = "@layer components{" + rule + "}";
        string html = "<style>" + rule + "</style><p><span class='chip'>" + (route == "pseudo" ? "" : "Status") + "</span></p>";
        if (route == "pseudo") {
            var document = HtmlDocumentParser.ParseDocument(html);
            HtmlComputedStyleSet styles = HtmlComputedStyleEngine.ComputeForContentSafety(document);
            Assert.True(styles.TryGetPseudoStyle(document.QuerySelector(".chip")!, HtmlPseudoElementKind.Before, out HtmlComputedStyle pseudo));
            Assert.Equal("4px", pseudo.GetValue("border-top-width"));
            Assert.Equal("solid", pseudo.GetValue("border-top-style"));
            Assert.Equal("green", pseudo.GetValue("border-top-color"));
            return;
        }
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderShape border = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Shape.StrokeWidth > 0D);
        Assert.Equal(4D, border.Shape.StrokeWidth);
        Assert.Equal(OfficeColor.Green, border.Shape.StrokeColor);
    }

    [Theory]
    [InlineData("border:2px solid red;border-color:var(--tone);border:4px solid green", 4D, false)]
    [InlineData("border:4px solid green;border-color:var(--tone)", 4D, true)]
    [InlineData("border-color:var(--tone);border:4px solid green", 4D, false)]
    [InlineData("border:2px solid red;border-color:var(--tone)!important;border:4px solid green", 4D, true)]
    [InlineData("border:4px solid green!important;border-color:var(--tone)", 4D, false)]
    [InlineData("border-color:var(--tone)!important;border:4px solid green!important", 4D, false)]
    [InlineData("border:4px solid green!important;border-color:var(--tone)!important", 4D, true)]
    [InlineData("border:4px solid green;border-width:var(--wide);border-width:6px", 6D, false)]
    [InlineData("border:4px solid green;border-width:6px;border-width:var(--wide)", 2D, false)]
    [InlineData("border:4px solid green;border:not-a-border", 4D, false)]
    [InlineData("border:4px solid green;border-width:-1px", 4D, false)]
    public void HtmlBorders_AuthoredOrderAndImportanceApplyAcrossOrdinaryAndVariableDeclarations(string declarations, double width, bool blue) {
        string html = "<style>.chip{display:inline-block;--tone:blue;--wide:2px;" + declarations
            + "}</style><p><span class='chip'>Status</span></p>";
        HtmlRenderShape border = Assert.Single(HtmlRenderTestDriver.Render(html).Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.chip" && shape.Shape.StrokeWidth > 0D);
        Assert.Equal(width, border.Shape.StrokeWidth);
        Assert.Equal(blue ? OfficeColor.Blue : OfficeColor.Green, border.Shape.StrokeColor);
    }

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
