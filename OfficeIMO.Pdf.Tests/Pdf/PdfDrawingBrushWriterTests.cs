using System;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingBrushWriterTests {
    [Fact]
    public void AlphaMaskAndGroupOpacityAreSerializedAsSeparateIsolatedGroups() {
        var shape = OfficeShape.Rectangle(40, 40); shape.FillColor = OfficeColor.Red;
        var content = new OfficeDrawing(60, 60).AddShape(shape, 5, 5).AddShape(shape, 5, 5);
        var maskShape = OfficeShape.Rectangle(60, 60); maskShape.FillColor = OfficeColor.FromRgba(0, 0, 0, 128);
        var mask = new OfficeDrawingSoftMask(new OfficeDrawing(60, 60).AddShape(maskShape, 0, 0));
        var drawing = new OfficeDrawing(60, 60).AddEffectDrawing(content, OfficeTransform.Identity, OfficeBlendMode.Normal, mask, 0.5);
        string pdf = Encoding.ASCII.GetString(Write(drawing));
        // Mask paint, masked content, and final opacity composite have distinct
        // transparency groups. Native rendering of this overlap is qualified separately.
        Assert.Equal(3, Regex.Matches(pdf, @"/Group << /S /Transparency /I true /K false >>").Count);
        Assert.Contains("/ca 1 /CA 1 /BM /Normal /SMask << /S /Alpha", pdf);
        Assert.Contains("/ca 0.5 /CA 0.5", pdf);
    }

    [Fact]
    public void PatternOpacityIsSerializedOnAnIsolatedComposite() {
        var shape = OfficeShape.Rectangle(20, 20); shape.FillColor = OfficeColor.Blue;
        var tile = new OfficeDrawing(20, 20).AddShape(shape, 0, 0);
        var drawing = new OfficeDrawing(60, 60).AddTilingPattern(tile, new OfficeImagePlacement(5, 5, 50, 50), 10, 10, opacity: 0.5);
        byte[] bytes = Write(drawing);
        string pdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/Group << /S /Transparency /I true /K false >>", pdf);
        Assert.Contains("/ca 0.5 /CA 0.5", pdf);
        Assert.Contains("/PatternType 1", pdf);
        Assert.Contains("/BBox [5 5 55 55]", pdf);
        var raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), background: OfficeColor.White);
        Assert.Equal(OfficeColor.White, raster.GetPixel(2, 2));
        OfficeColor overlap = raster.GetPixel(20, 20);
        Assert.InRange(overlap.R, 120, 135);
        Assert.InRange(overlap.G, 120, 135);
        Assert.Equal(255, overlap.B);
    }

    [Fact]
    public void TileOverflowCannotPaintIntoPatternGaps() {
        var shape = OfficeShape.Rectangle(20, 10); shape.FillColor = OfficeColor.Red;
        var overflow = new OfficeDrawing(20, 10).AddShape(shape, 0, 0);
        var tile = new OfficeDrawing(10, 10).AddEffectDrawing(overflow, OfficeTransform.Identity);
        var drawing = new OfficeDrawing(40, 10).AddTilingPattern(tile, new OfficeImagePlacement(0, 0, 40, 10), 20, 10);
        var raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(Write(drawing)), background: OfficeColor.White);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(5, 5));
        Assert.Equal(OfficeColor.White, raster.GetPixel(15, 5));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 5));
        Assert.Equal(OfficeColor.White, raster.GetPixel(35, 5));
    }

    [Fact]
    public void NestedPatternsRetainReusableCellsWithoutExpandingEveryTile() {
        var shape = OfficeShape.Rectangle(1, 1); shape.FillColor = OfficeColor.Red;
        var unit = new OfficeDrawing(1, 1).AddShape(shape, 0, 0);
        var tile = new OfficeDrawing(10, 1).AddTilingPattern(unit, new OfficeImagePlacement(0, 0, 10, 1), 1, 1, maximumTileCount: 16);
        var drawing = new OfficeDrawing(100, 1).AddTilingPattern(tile, new OfficeImagePlacement(0, 0, 100, 1), 10, 1, maximumTileCount: 16);
        byte[] bytes = Write(drawing);
        Assert.Contains("/PatternType 1", Encoding.ASCII.GetString(bytes));
        var raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes), background: OfficeColor.White);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(0, 0));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(50, 0));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(99, 0));
    }

    [Fact]
    public void UnsupportedMaskInterpretationsFailExplicitly() {
        var drawing = new OfficeDrawing(10, 10).AddEffectDrawing(new OfficeDrawing(10, 10), OfficeTransform.Identity,
            OfficeBlendMode.Normal, new OfficeDrawingSoftMask(new OfficeDrawing(10, 10), OfficeSoftMaskMode.Luminosity));
        Assert.Throws<NotSupportedException>(() => Write(drawing));
    }

    private static byte[] Write(OfficeDrawing drawing) => PdfDocument.Create(new PdfCreateOptions {
        PageWidth = global::OfficeIMO.OfficeLength.Points(drawing.Width), PageHeight = global::OfficeIMO.OfficeLength.Points(drawing.Height), CompressContentStreams = false,
        MarginLeft = global::OfficeIMO.OfficeLength.Points(0), MarginTop = global::OfficeIMO.OfficeLength.Points(0), MarginRight = global::OfficeIMO.OfficeLength.Points(0), MarginBottom = global::OfficeIMO.OfficeLength.Points(0)
    }).Drawing(drawing).ToBytes();

}
