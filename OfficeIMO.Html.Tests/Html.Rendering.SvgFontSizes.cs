using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlSvgFontSizeProjectionTests {
    [Theory]
    [InlineData("2em", "50%", 20D)]
    [InlineData("150%", "1em", 30D)]
    [InlineData("1rem", "2em", 32D)]
    public void RelativeSvgTextSizesUseComputedCssContext(string svgSize, string textSize, double expectedSize) {
        var source = HtmlConversionDocument.Parse("<style>html{font-size:16px}body{font-size:20px}svg{font-size:" + svgSize + "}text{font-size:" + textSize + "}</style>"
            + "<svg width='400' height='100'><text x='10' y='70'>Visible SVG</text></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        Assert.DoesNotContain(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(OfficeIMO.Tests.DrawingTestTraversal.Elements(drawing.Drawing).OfType<OfficeDrawingText>());
        Assert.Equal(expectedSize, text.Font.Size, 3);
    }

    [Fact]
    public void InheritedRelativeSizeDoesNotDegradeTextlessVectorContent() {
        var source = HtmlConversionDocument.Parse("<style>body{font-size:100%}</style><svg width='30' height='30'><rect width='30' height='30' fill='red'/></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        Assert.False(rendered.HasLoss);
        Assert.DoesNotContain(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
    }

    [Theory]
    [InlineData("", "font-size='24'", "font-size='12'", 12D)]
    [InlineData("text{font-size:50%}", "font-size='24'", "font-size='10'", 12D)]
    [InlineData("g{font-size:2em}text{font-size:50%}", "font-size='24'", "", 24D)]
    [InlineData("text{font-size:inherit}", "font-size='24'", "font-size='10'", 24D)]
    public void SvgPresentationSizesEstablishTheContextForCss(string css, string svgAttributes, string textAttributes, double expectedSize) {
        var source = HtmlConversionDocument.Parse("<style>body{font-size:20px}" + css + "</style>"
            + "<svg width='400' height='100' " + svgAttributes + "><g><text x='10' y='70' " + textAttributes + ">Visible SVG</text></g></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        Assert.DoesNotContain(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(OfficeIMO.Tests.DrawingTestTraversal.Elements(drawing.Drawing).OfType<OfficeDrawingText>());
        Assert.Equal(expectedSize, text.Font.Size, 3);
    }

    [Fact]
    public void RelativeSvgSizesRespectTheCallerDefaultFontSize() {
        var source = HtmlConversionDocument.Parse("<style>svg{font-size:2em}text{font-size:50%}</style>"
            + "<svg width='400' height='100'><text x='10' y='70'>Visible SVG</text></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions { DefaultFontSize = 24D });
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(OfficeIMO.Tests.DrawingTestTraversal.Elements(drawing.Drawing).OfType<OfficeDrawingText>());
        Assert.Equal(24D, text.Font.Size, 3);
    }

    [Theory]
    [InlineData("body{font-size:20px}g{font-size:initial}text{font-size:50%}", "font-size='24'", "", 8D)]
    [InlineData("body{font-size:20px}text{font-size:initial}", "font-size='24'", "font-size='10'", 16D)]
    [InlineData("text{font-size:inherit}", "font-size='24'", "font-size='10'", 24D)]
    [InlineData("text{font-size:unset}", "font-size='24'", "font-size='10'", 24D)]
    public void CssWideSvgFontSizesOverrideAttributes(string css, string groupAttributes, string textAttributes, double expectedSize) {
        var source = HtmlConversionDocument.Parse("<style>" + css + "</style><svg width='400' height='100'><g "
            + groupAttributes + "><text x='10' y='70' " + textAttributes + ">Visible SVG</text></g></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        Assert.DoesNotContain(rendered.Diagnostics, x => x.Code == HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(OfficeIMO.Tests.DrawingTestTraversal.Elements(drawing.Drawing).OfType<OfficeDrawingText>());
        Assert.Equal(expectedSize, text.Font.Size, 3);
    }

    [Fact]
    public void SvgRootInitialFontSizeUsesTheInitialSizeInsteadOfTheHtmlParentSize() {
        var source = HtmlConversionDocument.Parse("<style>body{font-size:20px}svg{font-size:initial}</style>"
            + "<svg width='400' height='100'><text x='10' y='70'>Visible SVG</text></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(OfficeIMO.Tests.DrawingTestTraversal.Elements(drawing.Drawing).OfType<OfficeDrawingText>());
        Assert.Equal(16D, text.Font.Size, 3);
    }

    [Fact]
    public void FontSizeResetAgreesAcrossComputedContainerContextAndRendering() {
        var source = HtmlConversionDocument.Parse("<style>body{font-size:24px}#reset{font-size:initial;width:10em;container-type:inline-size}"
            + "#probe{font-size:2em;color:blue}@container(min-width:200px){#probe{color:red}}</style>"
            + "<div id='reset'><span id='probe'>Child</span></div>");
        var styles = HtmlComputedStyleEngine.Compute(source);
        Assert.Equal(12D, styles[source.Document.QuerySelector("#reset")!].ResolvedFontSizePoints);
        Assert.Equal(24D, styles[source.Document.QuerySelector("#probe")!].ResolvedFontSizePoints);
        Assert.Equal("rgba(0, 0, 255, 1)", styles[source.Document.QuerySelector("#probe")!].GetValue("color"));
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions());
        var text = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals.OfType<HtmlRenderText>()), item => item.Text == "Child");
        Assert.Equal(32D, text.Font.Size, 3);
    }
}
