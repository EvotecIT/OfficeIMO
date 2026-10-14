using System.Collections.Generic;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioSvgNonPaintingDefinitionTests {
    [Fact]
    public void InlineClipPathContributesClippingGeometryWithoutPaintingIt() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>" +
                           "<rect width='20' height='10' fill='red' clip-path='url(#left)'/>" +
                           "<clipPath id='left'><rect width='10' height='20'/></clipPath></svg>";
        OfficeRasterImage raster = Rasterize(svg);

        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(15, 5));
        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 15));
    }

    [Theory]
    [InlineData("<mask id='unused'><rect width='10' height='20'/></mask>")]
    [InlineData("<marker id='unused'><rect width='10' height='20'/></marker>")]
    [InlineData("<pattern id='unused' width='10' height='20' patternUnits='userSpaceOnUse'><rect width='10' height='20'/></pattern>")]
    [InlineData("<filter id='unused'><feFlood flood-color='black'/></filter>")]
    public void UnreferencedInlineDefinitionsDoNotPaintOrReportVisualLoss(string definition) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>" +
                     "<rect x='10' width='10' height='20' fill='red'/>" + definition + "</svg>";
        OfficeRasterImage raster = Rasterize(svg);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 10));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 10));
    }

    [Fact]
    public void InlineSymbolOnlyPaintsWhenInstantiatedByUse() {
        const string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>" +
                           "<symbol id='badge' viewBox='0 0 10 20'><rect width='10' height='20' fill='red'/></symbol>" +
                           "<use href='#badge' x='10' width='10' height='20'/></svg>";
        OfficeRasterImage raster = Rasterize(svg);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 10));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 10));
    }

    [Theory]
    [InlineData("linearGradient")]
    [InlineData("radialGradient")]
    public void InlineGradientsRemainAvailableAsPaintServers(string elementName) {
        string svg = "<svg xmlns='http://www.w3.org/2000/svg' width='20' height='20'>" +
                     "<" + elementName + " id='accent'><stop offset='0' stop-color='red'/>" +
                     "<stop offset='1' stop-color='red'/></" + elementName + ">" +
                     "<rect x='10' width='10' height='20' fill='url(#accent)'/></svg>";
        OfficeRasterImage raster = Rasterize(svg);

        Assert.Equal(OfficeColor.Transparent, raster.GetPixel(5, 10));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(15, 10));
    }

    private static OfficeRasterImage Rasterize(string svg) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        Assert.True(VisioSvgPreviewRasterizer.TryRasterize(
            Encoding.UTF8.GetBytes(svg), null, null, null, null, null,
            diagnostics, "inline-definitions.svg", default, out OfficeRasterImage? image));
        Assert.Empty(diagnostics);
        return Assert.IsType<OfficeRasterImage>(image);
    }
}
