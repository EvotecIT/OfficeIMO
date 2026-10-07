using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Theory]
    [InlineData("calc(20px + 10%)", 40D)]
    [InlineData("min(50px, 30%)", 50D)]
    [InlineData("max(20px, 10%)", 20D)]
    [InlineData("clamp(20px, 40%, 60px)", 60D)]
    [InlineData("calc((8px + 2px) * 3)", 30D)]
    [InlineData("calc(90px / 3)", 30D)]
    [InlineData("calc((10px / 2px) * 5px)", 25D)]
    public void HtmlCssLengthMath_ResolvesAcrossSharedLayoutConsumers(string width, double expected) {
        string html = "<div id='math' style='width:" + width + ";height:10px;margin:0;background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D),
            BackgroundColor = OfficeColor.Transparent
        });

        HtmlRenderShape shape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#math");
        Assert.Equal(expected, shape.Width, 3);
    }

    [Theory]
    [InlineData("calc(10px + 2)")]
    [InlineData("calc(10px * 2px)")]
    [InlineData("calc(10px / 0)")]
    [InlineData("min(10px, 2)")]
    [InlineData("clamp(10px, 20px)")]
    [InlineData("calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(calc(1px)))))))))))))))))))))))))))))))))")]
    public void HtmlCssLengthMath_RejectsInvalidDimensionsAndUnboundedNesting(string width) {
        string html = "<div id='math-invalid' style='width:" + width + ";height:10px;margin:0;background:#ff0000'></div>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            ViewportHeight = 30D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape shape = Assert.Single(
            rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            item => item.Source == "div#math-invalid");
        Assert.NotEqual(10D, shape.Width);
    }

    [Fact]
    public void HtmlCssLengthMath_ClampsCalculatedSizingAtTheUsedValueBoundary() {
        const string html = "<div id='parent' style='width:200px'>"
            + "<div id='calculated' style='width:calc(10px - 20px)'></div>"
            + "<div id='literal' style='width:-10px'></div></div>";
        AngleSharp.Html.Dom.IHtmlDocument document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed,
            new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions(), new HtmlDiagnosticReport());
        HtmlRenderBoxStyle parent = resolver.Resolve(document.QuerySelector("#parent")!, 200D);

        Assert.Equal(0D, resolver.Resolve(document.QuerySelector("#calculated")!, 200D, parent).ExplicitWidth);
        Assert.Null(resolver.Resolve(document.QuerySelector("#literal")!, 200D, parent).ExplicitWidth);
    }

    [Fact]
    public void HtmlCssLengthMath_ResolvesVerticalPercentagesAgainstDefiniteParentHeight() {
        const string html = "<div id='parent' style='width:200px;height:100px'>"
            + "<div id='child' style='height:calc(10px + 50%);min-height:calc(5px + 25%);max-height:calc(30px + 50%)'></div>"
            + "</div><div id='indefinite'><div id='unresolved' style='height:calc(10px + 50%)'></div></div>";
        AngleSharp.Html.Dom.IHtmlDocument document = HtmlConversionDocument.Parse(html).CreateDocumentForRendering();
        IReadOnlyDictionary<AngleSharp.Dom.IElement, HtmlComputedStyle> computed = HtmlComputedStyleEngine.Compute(document);
        var styles = new HtmlComputedStyleSet(computed,
            new Dictionary<AngleSharp.Dom.IElement, HtmlPseudoElementStylePair>());
        var resolver = new HtmlRenderStyleResolver(styles, new HtmlRenderOptions(), new HtmlDiagnosticReport());
        HtmlRenderBoxStyle parent = resolver.Resolve(document.QuerySelector("#parent")!, 200D);
        HtmlRenderBoxStyle child = resolver.Resolve(document.QuerySelector("#child")!, 200D, parent);
        HtmlRenderBoxStyle indefinite = resolver.Resolve(document.QuerySelector("#indefinite")!, 200D);

        Assert.Equal(60D, child.ExplicitHeight);
        Assert.Equal(30D, child.MinHeight);
        Assert.Equal(80D, child.MaxHeight);
        Assert.Null(resolver.Resolve(document.QuerySelector("#unresolved")!, 200D, indefinite).ExplicitHeight);
    }

    [Fact]
    public void HtmlCssLengthMath_ClampsCalculatedContainerDimensionsBeforeQueries() {
        const string html = "<style>"
            + "@container (width:0px){#width-item{background:red}}"
            + "@container (width:0px){#max-width-item{background:red}}"
            + "@container (height:0px){#height-item{background:red}}"
            + "@container (height:0px){#max-height-item{background:red}}"
            + "</style>"
            + "<section style='width:calc(10px - 20px);container-type:inline-size'><div id='width-item' style='width:10px;height:10px;background:blue'></div></section>"
            + "<section style='width:100px;max-width:calc(10px - 20px);container-type:inline-size'><div id='max-width-item' style='width:10px;height:10px;background:blue'></div></section>"
            + "<section style='width:100px;height:calc(10px - 20px);container-type:size'><div id='height-item' style='width:10px;height:10px;background:blue'></div></section>"
            + "<section style='width:100px;height:100px;max-height:calc(10px - 20px);container-type:size'><div id='max-height-item' style='width:10px;height:10px;background:blue'></div></section>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { ViewportWidth = 300D, ViewportHeight = 300D });

        foreach (string id in new[] { "width-item", "max-width-item", "height-item", "max-height-item" }) {
            HtmlRenderShape shape = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderShape>(),
                item => item.Source == "div#" + id && item.Shape.FillColor.HasValue);
            Assert.Equal(OfficeColor.Red, shape.Shape.FillColor);
        }
    }
}
