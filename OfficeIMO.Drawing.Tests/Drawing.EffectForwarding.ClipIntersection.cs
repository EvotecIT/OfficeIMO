using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingEffectForwardingContractTests {
    [Theory]
    [InlineData(7L, false)]
    [InlineData(8L, true)]
    public void ChildTextProfileRetainsClipMaskAllocationLimit(long maximumPixels, bool succeeds) {
        // The output needs four pixels; the clip's color surface and alpha mask
        // together charge eight before any other intermediate has recorded a limit.
        var child = new OfficeDrawing(2, 2) { TextShapingLanguage = "pl" };
        OfficeShape rectangle = OfficeShape.Rectangle(2, 2);
        rectangle.FillColor = OfficeColor.Red;
        rectangle.StrokeColor = null;
        child.AddShape(rectangle, 0, 0);
        OfficeClipPath clip = OfficeClipPath.Path(OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.LineTo(2, 0), OfficePathCommand.LineTo(0, 2), OfficePathCommand.Close());
        var drawing = new OfficeDrawing(2, 2).AddClippedDrawing(child, 0, 0, clip);
        var options = new OfficeDrawingRasterRenderOptions {
            PdfRasterConventions = true,
            MaximumRasterPixels = maximumPixels
        };

        if (succeeds) {
            OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing, options);
            Assert.Equal(OfficeColor.Red, actual.GetPixel(0, 0));
            Assert.Equal(OfficeColor.FromRgba(255, 0, 0, 128), actual.GetPixel(1, 0));
        } else {
            Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing, options));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RepeatedNonrectangularClipCoversEffectPaintOnce(bool translatedImage) {
        var source = new OfficeDrawing(2, 1);
        AddFill(source, OfficeColor.Red);
        AddFill(source, OfficeColor.Blue);
        source.AddImageWithInterpolation(OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Black)),
            "image/png", new OfficeImageProjection(new OfficeImagePlacement(1, 0, 1, 1)), interpolate: false);
        OfficeRasterImage flattened = OfficeDrawingRasterRenderer.Render(source);
        double offset = translatedImage ? .5D : 0D;
        var effect = new OfficeDrawing(5, 2);
        if (translatedImage) {
            var imageLayer = new OfficeDrawing(4, 2).AddImageWithInterpolation(OfficePngWriter.Encode(flattened),
                "image/png", new OfficeImageProjection(new OfficeImagePlacement(0, 0, 4, 2)), interpolate: false);
            effect.AddEffectDrawing(imageLayer, OfficeTransform.Translate(offset, 0));
        } else {
            effect.AddEffectDrawing(source, OfficeTransform.Scale(2, 2));
        }
        var triangle = new[] { new OfficePoint(0, 0), new OfficePoint(4, 0), new OfficePoint(0, 2) };
        OfficeClipPath clip = OfficeClipPath.Path(OfficePathCommand.MoveTo(0, 0),
            OfficePathCommand.LineTo(4, 0), OfficePathCommand.LineTo(0, 2), OfficePathCommand.Close());
        var inner = new OfficeDrawing(5, 2).AddClippedDrawing(effect, 0, 0, clip);
        var drawing = new OfficeDrawing(5, 2).AddClippedDrawing(inner, 0, 0, clip);

        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { PdfRasterConventions = true });
        var expected = new OfficeRasterImage(5, 2);
        var canvas = new OfficeRasterCanvas(expected) { PdfRasterConventions = true };
        // An identical geometric clip is idempotent. Apply it once to independently
        // flattened paint, rather than multiplying masks or clipping each child.
        using (canvas.PushClipPolygon(triangle)) {
            if (translatedImage) {
                // A forwarded image retains the canonical PDF image placement and
                // sampling rules; an affine group bitmap uses a different grid.
                canvas.DrawImage(flattened,
                    new OfficeImageProjection(new OfficeImagePlacement(offset, 0, 4, 2)), interpolate: false);
            } else {
                canvas.DrawAffineImage(flattened, OfficeTransform.Scale(2, 2), 1D, interpolate: false);
            }
        }
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
        Assert.Equal(OfficeColor.Blue, actual.GetPixel(1, 0));
        if (translatedImage) Assert.Equal(0, actual.GetPixel(0, 0).A);
        else Assert.Equal(OfficeColor.FromRgba(0, 0, 255, 191), actual.GetPixel(0, 1));
    }
}
