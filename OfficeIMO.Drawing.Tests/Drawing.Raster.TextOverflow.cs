using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingRasterTests {
    [Fact]
    public void StyledPositionedTextUsesResolvedAdvanceWidth() {
        static OfficeRasterImage Render(double advance) {
            var drawing = new OfficeDrawing(160D, 50D).AddPositionedText(
                "Advance",
                4D,
                8D,
                145D,
                36D,
                new OfficeFontInfo("Arial", 24D, OfficeFontStyle.Bold),
                OfficeColor.Black,
                OfficeTextAlignment.Left,
                lineHeight: null,
                textAdvanceWidth: advance,
                underlineStyle: OfficeTextDecorationStyle.Wavy,
                strikethroughStyle: OfficeTextDecorationStyle.Double,
                baseline: OfficeTextBaseline.Superscript);
            return OfficeDrawingRasterRenderer.Render(drawing);
        }

        static int LastPaintedColumn(OfficeRasterImage image) {
            byte[] pixels = image.GetPixels();
            for (int x = image.Width - 1; x >= 0; x--) {
                for (int y = 0; y < image.Height; y++) {
                    if (pixels[(((y * image.Width) + x) * 4) + 3] != 0) return x;
                }
            }
            return -1;
        }

        int narrowRight = LastPaintedColumn(Render(36D));
        int wideRight = LastPaintedColumn(Render(112D));

        Assert.True(narrowRight >= 0);
        Assert.True(wideRight > narrowRight + 50, $"Expected the styled positioned advance to widen raster output, got {narrowRight} and {wideRight}.");
    }

    [Fact]
    public void JustifiedScriptWhitespaceBackgroundUsesRenderedBaseline() {
        var drawing = new OfficeDrawing(200D, 70D).AddRichText(
            new[] {
                new OfficeRichTextRun("left right\n", 20D, OfficeColor.Black, backgroundColor: OfficeColor.Yellow, baseline: OfficeTextBaseline.Superscript),
                new OfficeRichTextRun("end", 20D, OfficeColor.Black)
            },
            4D,
            4D,
            180D,
            60D,
            OfficeTextAlignment.Justify,
            lineHeight: 24D,
            wrapText: true);
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing);

        static (int Min, int Max) YellowBounds(OfficeRasterImage raster, int x) {
            int min = int.MaxValue;
            int max = -1;
            for (int y = 0; y < raster.Height; y++) {
                if (raster.GetPixel(x, y) == OfficeColor.Yellow) {
                    min = Math.Min(min, y);
                    max = y;
                }
            }
            return (min, max);
        }

        (int wordMin, int wordMax) = YellowBounds(image, 8);
        (int gapMin, int gapMax) = YellowBounds(image, 90);
        Assert.True(wordMax >= wordMin, "Expected a highlighted superscript word.");
        Assert.True(gapMax >= gapMin, "Expected the expanded justified space to retain its highlight.");
        Assert.Equal((wordMin, wordMax), (gapMin, gapMax));
    }

    [Fact]
    public void OfficeRasterCanvas_ClipOverflowKeepsExactTextWithoutImplicitInset() {
        const string text = "Sales";
        const double fontSize = 20D;
        var measuredImage = new OfficeRasterImage(180, 40, OfficeColor.Transparent);
        var measuredCanvas = new OfficeRasterCanvas(measuredImage);
        double advance = measuredCanvas.MeasureText(text, fontSize, "Arial");

        var positioned = new OfficeDrawing(180D, 40D).AddPositionedText(
            text,
            3D,
            0D,
            advance,
            32D,
            new OfficeFontInfo("Arial", fontSize),
            OfficeColor.Black);
        var bounded = new OfficeDrawing(180D, 40D).AddText(
            text,
            0D,
            0D,
            advance + 6D,
            32D,
            new OfficeFontInfo("Arial", fontSize),
            OfficeColor.Black);

        AssertRasterImagesEqual(
            OfficeDrawingRasterRenderer.Render(bounded),
            OfficeDrawingRasterRenderer.Render(positioned));
    }

    [Fact]
    public void OfficeDrawing_ClippedPositionedTextRetainsSourceAdvance() {
        const string text = "Quarterly Operations Dashboard";
        var drawing = new OfficeDrawing(180D, 40D).AddClippedPositionedText(
            text,
            0D,
            0D,
            160D,
            32D,
            0D,
            0D,
            OfficeClipPath.Rectangle(120D, 32D),
            new OfficeFontInfo("Arial", 20D),
            OfficeColor.Black,
            textAdvanceWidth: 150D);

        OfficeDrawingGroup group = Assert.IsType<OfficeDrawingGroup>(Assert.Single(drawing.Elements));
        OfficeDrawingText positioned = Assert.IsType<OfficeDrawingText>(Assert.Single(group.Drawing.Elements));
        Assert.Equal(text, positioned.Text);
        Assert.Equal(OfficeTextOverflowBehavior.Clip, positioned.OverflowBehavior);
        Assert.Equal(150D, positioned.TextAdvanceWidth);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.True(rendered.GetPixels().Any(static value => value != 0));
    }
}
