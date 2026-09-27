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
        var rendered = HtmlRenderEngine.Render(source,new HtmlRenderOptions());
        Assert.DoesNotContain(rendered.Diagnostics,x=>x.Code==HtmlRenderDiagnosticCodes.SvgContentUnsupported);
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(drawing.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(expectedSize,text.Font.Size,3);
    }

    [Fact]
    public void InheritedRelativeSizeDoesNotDegradeTextlessVectorContent() {
        var source = HtmlConversionDocument.Parse("<style>body{font-size:100%}</style><svg width='30' height='30'><rect width='30' height='30' fill='red'/></svg>");
        var rendered = HtmlRenderEngine.Render(source,new HtmlRenderOptions());
        Assert.False(rendered.HasLoss);
        Assert.DoesNotContain(rendered.Diagnostics,x=>x.Code==HtmlRenderDiagnosticCodes.SvgContentUnsupported);
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
        var text = Assert.Single(drawing.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(expectedSize, text.Font.Size, 3);
    }

    [Fact]
    public void RelativeSvgSizesRespectTheCallerDefaultFontSize() {
        var source = HtmlConversionDocument.Parse("<style>svg{font-size:2em}text{font-size:50%}</style>"
            + "<svg width='400' height='100'><text x='10' y='70'>Visible SVG</text></svg>");
        var rendered = HtmlRenderEngine.Render(source, new HtmlRenderOptions { DefaultFontSize = 24D });
        var drawing = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderDrawing>());
        var text = Assert.Single(drawing.Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.Equal(24D, text.Font.Size, 3);
    }
}
