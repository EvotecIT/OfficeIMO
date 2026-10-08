using System;
using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOrdinaryGroupOpacityTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void IsolatedOpacityAppliesOnceAfterOverlappingPaint(bool childAlpha, bool nested) {
        var children = new OfficeDrawing(100, 80);
        var first = OfficeShape.Rectangle(60, 60); first.FillColor = OfficeColor.Red; first.StrokeWidth = 0;
        first.FillOpacity = childAlpha ? .5 : 1;
        children.AddShape(first, 10, 10).AddShape(first.Clone(), 30, 10);
        var group = new OfficeDrawing(100, 80).AddEffectDrawing(children, OfficeTransform.Identity, .5);
        var scene = nested ? new OfficeDrawing(100, 80).AddEffectDrawing(group, OfficeTransform.Identity, .5) : group;
        byte[] pdf = Export(scene);
        var actual = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        var expected = OfficeDrawingRasterRenderer.Render(scene, background: OfficeColor.White);
        foreach (var point in new[] { (20, 35), (50, 35), (80, 35), (95, 35) })
            EqualPixel(expected.GetPixel(point.Item1, point.Item2), actual.GetPixel(point.Item1, point.Item2));
    }

    [Fact]
    public void ImageAndTextPaintRemainInsideTheGroupAndLogicalTextIsRetained() {
        var image = new OfficeRasterImage(2, 2, OfficeColor.Blue);
        byte[] png = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png);
        var children = new OfficeDrawing(100, 80);
        children.AddImage(png, "image/png", new OfficeImageProjection(new OfficeImagePlacement(10, 10, 60, 60)));
        children.AddImage(png, "image/png", new OfficeImageProjection(new OfficeImagePlacement(30, 10, 60, 60)));
        children.AddText("Grouped", 10, 10, 80, 20, color: OfficeColor.White);
        var scene = new OfficeDrawing(100, 80).AddEffectDrawing(children, OfficeTransform.Identity, .5);
        byte[] pdf = Export(scene); var page = PdfReadDocument.Open(pdf).Pages[0];
        Assert.Contains("Grouped", page.ExtractText());
        var actual = OfficeDrawingRasterRenderer.Render(page.ToDrawing(), background: OfficeColor.White);
        foreach (var point in new[] { (20, 50), (50, 50), (80, 50) }) {
            var color = actual.GetPixel(point.Item1, point.Item2);
            Assert.InRange((int)color.R, 126, 130); Assert.InRange((int)color.G, 126, 130);
            Assert.Equal(255, color.B);
        }
        int textPixels = 0;
        for (int y = 10; y < 30; y++) for (int x = 10; x < 90; x++) {
            OfficeColor color = actual.GetPixel(x, y);
            if (color.R > 150 && color.G > 150 && color.B == 255) textPixels++;
        }
        Assert.True(textPixels > 20, "Contrasting text must be painted inside the translucent image group.");

    }

    [Fact]
    public void NestedMaskGroupsAndPageGroupsHaveIndependentPaintScopes() {
        var children = new OfficeDrawing(100, 80);
        var red = OfficeShape.Rectangle(60, 60); red.FillColor = OfficeColor.Red; red.StrokeWidth = 0;
        children.AddShape(red, 10, 10).AddShape(red.Clone(), 30, 10);
        var maskPaint = new OfficeDrawing(100, 80);
        var white = OfficeShape.Rectangle(80, 60); white.FillColor = OfficeColor.White; white.StrokeWidth = 0;
        maskPaint.AddShape(white, 10, 10);
        var mask = new OfficeDrawingSoftMask(new OfficeDrawing(100, 80)
            .AddEffectDrawing(maskPaint, OfficeTransform.Identity, .5));
        var scene = new OfficeDrawing(100, 80).AddEffectDrawing(children, OfficeTransform.Identity,
            OfficeBlendMode.Normal, mask, .5);
        var expected = OfficeDrawingRasterRenderer.Render(scene, background: OfficeColor.White);
        var actual = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(Export(scene)).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        foreach (var point in new[] { (20, 35), (50, 35), (80, 35), (95, 35) })
            EqualPixel(expected.GetPixel(point.Item1, point.Item2), actual.GetPixel(point.Item1, point.Item2));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void IndependentlyProducedCairoGroupsRetainCompositeOpacity(bool childAlpha, bool nested) {
        string name = "cairo-" + (childAlpha ? "child-alpha" : "opaque") + (nested ? "-nested" : "") + ".pdf";
        string path = Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability", "Transparency", name);
        var image = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0].ToDrawing(),
            background: OfficeColor.White);
        double groupAlpha = nested ? .25 : .5;
        double singleAlpha = groupAlpha * (childAlpha ? .5 : 1);
        double overlapAlpha = groupAlpha * (childAlpha ? .75 : 1);
        EqualPixel(OfficeColor.FromRgb(255, (byte)Math.Round(255*(1-singleAlpha)), (byte)Math.Round(255*(1-singleAlpha))), image.GetPixel(20,35));
        EqualPixel(OfficeColor.FromRgb(255, (byte)Math.Round(255*(1-overlapAlpha)), (byte)Math.Round(255*(1-overlapAlpha))), image.GetPixel(50,35));
        EqualPixel(OfficeColor.White, image.GetPixel(95,35));
    }

    private static byte[] Export(OfficeDrawing drawing) => PdfDocument.Create(new PdfOptions {
        PageWidth = drawing.Width, PageHeight = drawing.Height, MarginLeft = 0, MarginRight = 0, MarginTop = 0, MarginBottom = 0
    }).Compose(compose => compose.Page(page => page.Content(content => content.Drawing(drawing)))).ToBytes();

    private static void EqualPixel(OfficeColor expected, OfficeColor actual) {
        Assert.InRange(Math.Abs(expected.R - actual.R), 0, 3);
        Assert.InRange(Math.Abs(expected.G - actual.G), 0, 3);
        Assert.InRange(Math.Abs(expected.B - actual.B), 0, 3);
    }
}
