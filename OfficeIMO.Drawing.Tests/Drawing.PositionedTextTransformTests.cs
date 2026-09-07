using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPositionedTextTransformTests {
    [Theory]
    [InlineData(90, false, false)]
    [InlineData(180, false, false)]
    [InlineData(270, false, false)]
    [InlineData(90, true, false)]
    [InlineData(90, false, true)]
    public void TransformPreservesTheCompletePositionedGlyphImage(int rotation, bool mirror, bool clipped) {
        const int size = 100;
        var font = new OfficeFontInfo("Arial", 12D);
        var source = new OfficeDrawing(size, size).AddPositionedText("and and", 20, 30, 40, 20, font,
            OfficeColor.Black, textAdvanceWidth: 36);
        var target = new OfficeDrawing(size, size);
        var frame = new OfficeImageFrameTransform(rotation, 50, 50, mirror);
        if (clipped) target.AddClippedPositionedText("and and", 20, 30, 40, 20, 40, 20,
            OfficeClipPath.Rectangle(20, 60), frame, font, OfficeColor.Black, textAdvanceWidth: 36);
        else target.AddPositionedText("and and", 20, 30, 40, 20, frame, font, OfficeColor.Black, textAdvanceWidth: 36);
        OfficeRasterImage original = OfficeDrawingRasterRenderer.Render(source);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(target.Clone());
        int sourceInk = 0, differences = 0;
        for (int y = 0; y < size; y++) for (int x = 0; x < size; x++) {
            int mirroredX = mirror ? size - 1 - x : x;
            var destination = rotation switch {
                90 => (X: size - 1 - y, Y: mirroredX),
                180 => (X: size - 1 - mirroredX, Y: size - 1 - y),
                _ => (X: y, Y: size - 1 - mirroredX)
            };
            int alpha = original.GetPixel(x, y).A;
            if (clipped && (destination.X < 40 || destination.X >= 60 || destination.Y < 20 || destination.Y >= 80)) alpha = 0;
            if (alpha > 0) sourceInk++;
            if (Math.Abs(alpha - actual.GetPixel(destination.X, destination.Y).A) > 2) differences++;
        }
        Assert.True(sourceInk > 20);
        Assert.Equal(0, differences);
    }

    [Fact]
    public void TransformedClippedTextCannotAllocateAnUnboundedOffscreenLayer() {
        var drawing = new OfficeDrawing(10, 10).AddClippedPositionedText("and", 0, 0, 100000, 100000,
            0, 0, OfficeClipPath.Rectangle(10, 10), new OfficeImageFrameTransform(90, 0, 0));
        Assert.Throws<OfficeImageExportLimitException>(() => OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { MaximumRasterPixels = 10000 }));
    }
}
