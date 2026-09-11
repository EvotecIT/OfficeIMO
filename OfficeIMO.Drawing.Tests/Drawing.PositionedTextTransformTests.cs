using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPositionedTextTransformTests {
    [Theory]
    [InlineData(90, OfficeFontStyle.Underline, 40)]
    [InlineData(180, OfficeFontStyle.Strikethrough, 2)]
    [InlineData(270, OfficeFontStyle.Underline | OfficeFontStyle.Strikethrough, 2)]
    public void RotatedWhitespaceRetainsDecorationOutsideItsFrame(int rotation, OfficeFontStyle style, int height) {
        const int size = 200;
        var fonts = new OfficeFontFaceCollection().Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        var font = new OfficeFontInfo("Proof Sans", 40, style);
        var source = new OfficeDrawing(size, size).AddPositionedText(" ", 20, 30, 10, height, font, OfficeColor.Black, textAdvanceWidth: 100);
        var target = new OfficeDrawing(size, size).AddPositionedText(" ", 20, 30, 10, height,
            new OfficeImageFrameTransform(rotation, 100, 100), font, OfficeColor.Black, textAdvanceWidth: 100);
        source.Fonts.AddRange(fonts); target.Fonts.AddRange(fonts);
        OfficeRasterImage original = OfficeDrawingRasterRenderer.Render(source);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(target);
        int ink = 0;
        for (int y = 0; y < size; y++) for (int x = 0; x < size; x++) {
            var destination = rotation switch {
                90 => (X: size - 1 - y, Y: x),
                180 => (X: size - 1 - x, Y: size - 1 - y),
                _ => (X: y, Y: size - 1 - x)
            };
            int alpha = original.GetPixel(x, y).A;
            if (alpha > 0) ink++;
            Assert.InRange(Math.Abs(alpha - actual.GetPixel(destination.X, destination.Y).A), 0, 2);
        }
        Assert.True(ink > 50);
    }

    [Theory]
    [InlineData(90, OfficeFontStyle.Italic, "f")]
    [InlineData(180, OfficeFontStyle.Italic, "j")]
    [InlineData(270, OfficeFontStyle.Bold | OfficeFontStyle.Italic, "f")]
    public void RotatedPositionedGlyphRetainsInkOutsideAdvanceFrame(int rotation, OfficeFontStyle style, string value) {
        const int size = 160;
        var fonts = new OfficeFontFaceCollection().Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        var metrics = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: fonts);
        var font = new OfficeFontInfo("Proof Sans", 40, style);
        double advance = metrics.MeasureText(value, font.Size, font.FamilyName, font.Style);
        var source = new OfficeDrawing(size, size).AddPositionedText(value, 60, 50, advance, 50, font, OfficeColor.Black, textAdvanceWidth: advance);
        var target = new OfficeDrawing(size, size).AddPositionedText(value, 60, 50, advance, 50,
            new OfficeImageFrameTransform(rotation, 80, 80), font, OfficeColor.Black, textAdvanceWidth: advance);
        source.Fonts.AddRange(fonts); target.Fonts.AddRange(fonts);
        OfficeRasterImage original = OfficeDrawingRasterRenderer.Render(source);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(target);
        int overhang = 0;
        for (int y = 0; y < size; y++) for (int x = 0; x < size; x++) {
            var destination = rotation switch {
                90 => (X: size - 1 - y, Y: x),
                180 => (X: size - 1 - x, Y: size - 1 - y),
                _ => (X: y, Y: size - 1 - x)
            };
            int alpha = original.GetPixel(x, y).A;
            if ((x < 60 || x >= 60 + advance) && alpha > 0) overhang++;
            Assert.InRange(Math.Abs(alpha - actual.GetPixel(destination.X, destination.Y).A), 0, 2);
        }
        Assert.True(overhang > 0, "The fixture must exercise ink outside the advance frame.");
    }

    [Theory]
    [InlineData("\n")]
    [InlineData("\r")]
    [InlineData("\r\n")]
    public void OrdinaryTextPreservesHardLinePlacement(string separator) {
        var drawing = new OfficeDrawing(100, 100).AddText("AAAA" + separator + "A", 10, 10, 80, 60,
            new OfficeFontInfo("Proof Sans", 20), alignment: OfficeTextAlignment.Center, lineHeight: 25);
        drawing.Fonts.Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        var actual = OfficeDrawingRasterRenderer.Render(drawing);
        int firstLineInk = 0, secondLineInk = 0;
        for (int y = 10; y < 60; y++) for (int x = 10; x < 90; x++) {
            if (actual.GetPixel(x, y).A == 0) continue;
            if (y < 35) firstLineInk++; else secondLineInk++;
        }
        Assert.True(firstLineInk > 20);
        Assert.True(secondLineInk > 20);
        Assert.True(firstLineInk > secondLineInk * 2);
    }

    [Theory]
    [InlineData(0, OfficeTextAlignment.Center, "\n")]
    [InlineData(0, OfficeTextAlignment.Right, "\r")]
    [InlineData(90, OfficeTextAlignment.Center, "\r\n")]
    [InlineData(90, OfficeTextAlignment.Right, "\n")]
    public void PositionedHardLinesMatchIndependentlyPositionedLines(int rotation, OfficeTextAlignment alignment, string separator) {
        var fonts = new OfficeFontFaceCollection().Add("Proof Sans", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf")));
        var metrics = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: fonts);
        var font = new OfficeFontInfo("Proof Sans", 20);
        var frame = new OfficeImageFrameTransform(rotation, 50, 50);
        var actualDrawing = new OfficeDrawing(100, 100).AddPositionedText("AAAA" + separator + "A", 10, 10, 80, 60,
            frame, font, OfficeColor.Black, alignment: alignment, lineHeight: 25);
        actualDrawing.Fonts.AddRange(fonts);
        var expectedDrawing = new OfficeDrawing(100, 100);
        expectedDrawing.Fonts.AddRange(fonts);
        expectedDrawing.AddPositionedText("AAAA", 10, 10, 80, 25, frame, font, OfficeColor.Black,
            alignment: alignment, textAdvanceWidth: metrics.MeasureText("AAAA", 20, "Proof Sans"));
        expectedDrawing.AddPositionedText("A", 10, 35, 80, 25, frame, font, OfficeColor.Black,
            alignment: alignment, textAdvanceWidth: metrics.MeasureText("A", 20, "Proof Sans"));
        var actual = OfficeDrawingRasterRenderer.Render(actualDrawing);
        var expected = OfficeDrawingRasterRenderer.Render(expectedDrawing);
        int ink = 0;
        for (int y = 0; y < 100; y++) for (int x = 0; x < 100; x++) {
            if (expected.GetPixel(x, y).A > 0) ink++;
            Assert.InRange(Math.Abs(actual.GetPixel(x, y).A - expected.GetPixel(x, y).A), 0, 2);
        }
        Assert.True(ink > 20);
    }

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
