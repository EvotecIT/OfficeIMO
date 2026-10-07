using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPositionedTextCffStateTests {
    [Fact]
    public void TransformedRunPreservesTheSurroundingDirectTextSequence() {
        byte[] font = CffRandomGlyphTestAssets.CreateRandomOverhangFont();
        OfficeDrawing Create(bool transformed) {
            var drawing = new OfficeDrawing(1000, 140);
            drawing.Fonts.Add("Random CFF", font);
            var face = new OfficeFontInfo("Random CFF", 20);
            drawing.AddPositionedText("A", 200, 60, 20, 30, face, OfficeColor.Black, textAdvanceWidth: 20);
            if (transformed) drawing.AddPositionedText("A", 500, 60, 20, 30,
                new OfficeImageFrameTransform(0, 510, 75, flipHorizontal: true), face, OfficeColor.Black, textAdvanceWidth: 20);
            drawing.AddPositionedText("A", 800, 60, 20, 30, face, OfficeColor.Black, textAdvanceWidth: 20);
            return drawing;
        }
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(Create(false));
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(Create(true));
        int painted = 0;
        for (int y = 0; y < expected.Height; y++) for (int x = 650; x < expected.Width; x++) {
            OfficeColor pixel = expected.GetPixel(x, y);
            if (pixel.A > 0) painted++;
            Assert.Equal(pixel, actual.GetPixel(x, y));
        }
        Assert.True(painted > 0);
    }

    [Fact]
    public void TransformedMeasurementsAndPaintShareTheDocumentOperationLimit() {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        var drawing = new OfficeDrawing(100, 40);
        drawing.Fonts.Add("Source Sans", font);
        var cff = Assert.IsAssignableFrom<IOfficeCffBoundedFontProgram>(Assert.Single(drawing.Fonts.Faces).Program);
        var budget = new OfficeCffOperationBudget();
        Assert.NotEmpty(cff.GetTextContoursBounded("A", 0, 0, 1, 100_000, System.Threading.CancellationToken.None, budget));
        int operations = 1_000_000 - budget.RemainingOperations;
        Assert.True(operations > 0);
        const int runLength = 128;
        // Measurement alone fits the allowance; measurement plus actual paint exceeds it.
        int runCount = 1_000_000 / (2 * operations * runLength) + 1;
        Assert.True((long)runCount * operations * runLength < 1_000_000);
        for (int index = 0; index < runCount; index++) drawing.AddPositionedText(new string('A', runLength), 0, 0, 80, 20,
            new OfficeImageFrameTransform(0, 40, 10, flipHorizontal: true), new OfficeFontInfo("Source Sans", 1),
            OfficeColor.Black, textAdvanceWidth: 80);
        Assert.Throws<InvalidDataException>(() => OfficeDrawingRasterRenderer.Render(drawing));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void RepeatedFrameTransformedRunsMatchIndependentGlyphPaint(bool clipped, bool mirrored) {
        const int width = 1800, height = 140;
        byte[] font = CffRandomGlyphTestAssets.CreateRandomOverhangFont();
        OfficeDrawing Create() {
            var drawing = new OfficeDrawing(width, height);
            drawing.Fonts.Add("Random CFF", font);
            return drawing;
        }
        void Add(OfficeDrawing drawing, int index) {
            double x = (index + 1) * 200D;
            var face = new OfficeFontInfo("Random CFF", 20);
            var frame = new OfficeImageFrameTransform(mirrored ? 0 : 180, x + 10, 75, flipHorizontal: mirrored);
            if (clipped) drawing.AddClippedPositionedText("A", x, 60, 20, 30, 0, 0,
                OfficeClipPath.Rectangle(width, height), frame, face, OfficeColor.Black, textAdvanceWidth: 20);
            else drawing.AddPositionedText("A", x, 60, 20, 30, frame, face, OfficeColor.Black, textAdvanceWidth: 20);
        }
        var repeated = Create();
        var expected = new OfficeRasterImage(width, height);
        int painted = 0;
        for (int index = 0; index < 8; index++) {
            Add(repeated, index);
            var independent = Create();
            Add(independent, index);
            OfficeRasterImage single = OfficeDrawingRasterRenderer.Render(independent);
            for (int y = 0; y < height; y++) for (int x = 0; x < width; x++) {
                OfficeColor pixel = single.GetPixel(x, y);
                if (pixel.A == 0) continue;
                Assert.Equal(0, expected.GetPixel(x, y).A); // The independent runs do not overlap.
                expected.SetPixel(x, y, pixel);
                painted++;
            }
        }
        Assert.True(painted > 0);
        byte[] actual = OfficeDrawingRasterRenderer.Render(repeated).GetPixels();
        int difference = expected.GetPixels().Zip(actual, (left, right) => Math.Abs(left - right)).Max();
        Assert.True(difference == 0, "Repeated frame-transformed runs clipped glyph paint; maximum channel difference: " + difference);
    }
}
