using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Threading;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingPreparedTextShapingTests {
    private static byte[] TrackedDistinctOutlines() => ManagedTextShapingTestAssets.AddTrackingTable(
        ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B'),
        ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { -1D, 1D },
            new[] { new short[] { -120, -40 }, new short[] { -80, 40 } }));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CompleteClusterSlicesRetainContinuationOriginsAndTracking(bool negative) {
        var program = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(TrackedDistinctOutlines()));
        const string text = "AB\u0301D";
        OfficeShapedGlyph attached = new(2, "B\u0301", 1, advanceWidth: 0, offsetX: 71, offsetY: -113);
        OfficeShapedGlyph continuation = OfficeShapedGlyph.CreatePositionedUnicodeContinuation(
            1, 1, negative ? -600 : 600, 0, -24, 17);
        OfficeShapedGlyph[] glyphs = negative ? new[] {
            new OfficeShapedGlyph(1, "D", 3, advanceWidth: -400, offsetX: 30, offsetY: -20),
            attached, continuation, new OfficeShapedGlyph(1, "A", 0, advanceWidth: -500)
        } : new[] {
            new OfficeShapedGlyph(1, "A", 0), attached, continuation, new OfficeShapedGlyph(1, "D", 3)
        };
        var result = new OfficeTextShapingResult(glyphs,
            negative ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        using var cancellation = new CancellationTokenSource();
        var prepared = new OfficePreparedTextShaping(program, text, result, cancellation.Token);

        Assert.False(prepared.TrySlice(1, 1, out _)); // The mark belongs to its complete logical cluster.
        Assert.True(prepared.TrySlice(1, 2, out OfficePreparedTextShaping? cluster));
        Assert.True(prepared.TrySlice(0, 1, out OfficePreparedTextShaping? logicalA));
        Assert.True(prepared.TrySlice(3, 1, out OfficePreparedTextShaping? logicalD));
        Assert.Same(program, cluster!.Program);
        Assert.Equal("B\u0301", cluster.Text);
        Assert.Equal(2, cluster.Result.Glyphs.Count);
        Assert.Equal(new[] { "B\u0301", string.Empty }, cluster.Result.Glyphs.Select(glyph => glyph.UnicodeText));
        Assert.All(cluster.Result.Glyphs, glyph => Assert.Equal(0, glyph.TextIndex));
        Assert.Equal(0, cluster.Result.Glyphs[0].AdvanceWidth);
        Assert.Equal(negative ? -600 : 600, cluster.Result.Glyphs[1].AdvanceWidth);
        Assert.Equal((71, -113), (cluster.Result.Glyphs[0].OffsetX, cluster.Result.Glyphs[0].OffsetY));
        Assert.Equal((-24, 17), (cluster.Result.Glyphs[1].OffsetX, cluster.Result.Glyphs[1].OffsetY));
        if (!negative) Assert.Null(Assert.Single(logicalA!.Result.Glyphs).AdvanceWidth);
        foreach (double size in new[] { 10D, 20D }) {
            Assert.Equal(((IOfficeFontProgram)program).MeasureShapedText(text, result, size), prepared.Measure(size), 9);
            Assert.Equal(prepared.Measure(size), logicalA!.Measure(size) + cluster.Measure(size) + logicalD!.Measure(size), 9);
            var fullContours = prepared.GetContours(10D, 3D, size, 1000, CancellationToken.None);
            // A is the visual prefix for positive advances and the visual tail for negative advances.
            var slicedContours = cluster.GetContours(10D + logicalA.Measure(size), 3D, size, 1000, CancellationToken.None);
            Assert.Equal(4, fullContours.Count); // Every selected synthetic glyph has one distinct outline.
            Assert.Equal(2, slicedContours.Count);
            for (int contour = 0; contour < slicedContours.Count; contour++) {
                Assert.Equal(fullContours[contour + 1].Count, slicedContours[contour].Count);
                for (int point = 0; point < slicedContours[contour].Count; point++) {
                    Assert.Equal(fullContours[contour + 1][point].X, slicedContours[contour][point].X, 9);
                    Assert.Equal(fullContours[contour + 1][point].Y, slicedContours[contour][point].Y, 9);
                }
            }
        }

        var mixed = new OfficePreparedTextShaping(program, text, new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 1200), attached,
            OfficeShapedGlyph.CreatePositionedUnicodeContinuation(1, 1, -600, 0, -24, 17),
            new OfficeShapedGlyph(1, "D", 3, advanceWidth: 400)
        }, OfficeTextDirection.LeftToRight));
        Assert.False(mixed.TrySlice(1, 2, out _)); // Independent painting would change pen normalization.
        var stationary = new OfficePreparedTextShaping(program, text, new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 1200), attached,
            OfficeShapedGlyph.CreatePositionedUnicodeContinuation(1, 1, 0, 0, -24, 17),
            new OfficeShapedGlyph(1, "D", 3, advanceWidth: 400)
        }, OfficeTextDirection.LeftToRight));
        Assert.False(stationary.TrySlice(1, 2, out _)); // Zero design advance cannot hide signed tracking.
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => { prepared.TrySlice(1, 2, out _); });
        Assert.Throws<OperationCanceledException>(() => { cluster.TrySlice(0, 2, out _); });
        Assert.Throws<OperationCanceledException>(() => prepared.GetContours(0, 0, 10, 1000, cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PreparedRasterRetainsLogicalTrackingAtExportScale(bool explicitAdvances) {
        var fonts = new OfficeFontFaceCollection().Add("Prepared Tracked", TrackedDistinctOutlines());
        IOfficeFontProgram program = Assert.Single(fonts.Faces).Program;
        var result = new OfficeTextShapingResult(new[] {
            explicitAdvances ? new OfficeShapedGlyph(1, "A", 0, advanceWidth: 700) : new OfficeShapedGlyph(1, "A", 0),
            explicitAdvances ? new OfficeShapedGlyph(2, "B", 1, advanceWidth: 400) : new OfficeShapedGlyph(2, "B", 1)
        }, OfficeTextDirection.LeftToRight);
        var prepared = new OfficePreparedTextShaping(program, "AB", result);
        const double logicalSize = 15D;
        double advance = prepared.Measure(logicalSize);
        var actual = new OfficeDrawing(50, 30).AddPositionedText("AB", 2, 2, advance, 20,
            new OfficeFontInfo("Prepared Tracked", logicalSize), OfficeColor.Black, textAdvanceWidth: advance);
        actual.Fonts.AddRange(fonts);
        actual.MarkLastTextWithPreparedShaping(prepared);
        var expected = new OfficeDrawing(50, 30).AddPositionedText("AB", 2, 2, advance, 20,
            new OfficeFontInfo("Prepared Tracked", logicalSize), OfficeColor.Black, textAdvanceWidth: advance);
        expected.Fonts.AddRange(fonts);
        expected.TextShapingProvider = new RetainedRunProvider(result);
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(100, 60), font: null, fonts: fonts);
        using (canvas.PushPreparedShaping(prepared)) {
            canvas.FontMetricScale = 2D;
            Assert.Equal(advance * 2D, canvas.MeasureText("AB", logicalSize * 2D, "Prepared Tracked"), 9);
            canvas.FontMetricScale = 1D;
            Assert.Equal(advance, canvas.MeasureText("AB", logicalSize, "Prepared Tracked"), 9);
        }
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(actual, 2D);
        byte[] pixels = image.GetPixels();
        Assert.Contains(Enumerable.Range(0, image.Width * image.Height), index => {
            OfficeColor pixel = image.GetPixel(index % image.Width, index / image.Width);
            return pixel.A > 0 && pixel.R < 128 && pixel.G < 128 && pixel.B < 128;
        });
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected, 2D).GetPixels(), pixels);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicCompositionRetainsPreparedClusterInk(bool scaled) {
        OfficeDrawing source = PreparedCopySource(OfficeColor.Black, out OfficePreparedTextShaping prepared);
        var translated = new OfficeDrawing(120D, 100D).AddDrawing(source, 7D, 9D);
        OfficeDrawingText text = Assert.Single(translated.Elements.OfType<OfficeDrawingText>());
        Assert.Same(prepared, text.PreparedShaping);
        Assert.Same(prepared.Program, text.PreparedShaping!.Program);
        Assert.Equal("A", text.Text);
        Assert.Equal("A", text.RasterText);
        Assert.False(text.PreservesPaintedGlyphs); // Prepared logical text is not a painted glyph alias string.
        double scale = scaled ? 2D : 1D;
        OfficeDrawing composed = scaled
            ? new OfficeDrawing(240D, 200D).AddEffectDrawing(translated, OfficeTransform.Scale(scale, scale))
            : translated;
        OfficeDrawing expected = PreparedCopyReference(OfficeColor.Black, 7D, 9D, scale);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected, background: OfficeColor.White).GetPixels(),
            OfficeDrawingRasterRenderer.Render(composed, background: OfficeColor.White).GetPixels());
    }

    [Fact]
    public void ColorTintRetainsPreparedClusterInkAndSourceAlpha() {
        OfficeColor sourceColor = OfficeColor.FromRgba(0, 0, 0, 128);
        OfficeDrawing source = PreparedCopySource(sourceColor, out OfficePreparedTextShaping prepared);
        OfficeDrawing tinted = source.Clone();
        tinted.ApplyColorTint(OfficeColor.FromRgba(255, 0, 0, 7));
        OfficeDrawingText text = Assert.Single(tinted.Elements.OfType<OfficeDrawingText>());
        Assert.Same(prepared, text.PreparedShaping);
        Assert.Same(prepared.Result, text.PreparedShaping!.Result);
        Assert.Equal("A", text.Text);
        Assert.Equal("A", text.RasterText);
        Assert.False(text.PreservesPaintedGlyphs);
        OfficeColor expectedColor = OfficeColor.FromRgba(255, 0, 0, 128);
        Assert.Equal(expectedColor, text.Color);
        OfficeDrawing expected = PreparedCopyReference(expectedColor, 0D, 0D, 1D);
        Assert.Equal(OfficeDrawingRasterRenderer.Render(expected, background: OfficeColor.White).GetPixels(),
            OfficeDrawingRasterRenderer.Render(tinted, background: OfficeColor.White).GetPixels());
    }

    private static OfficeDrawing PreparedCopySource(OfficeColor color, out OfficePreparedTextShaping prepared) {
        const string family = "Prepared Copy";
        var drawing = new OfficeDrawing(100D, 80D);
        drawing.Fonts.Add(family, TrackedDistinctOutlines());
        IOfficeFontProgram program = Assert.Single(drawing.Fonts.Faces).Program;
        var result = new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(2, "A", 0, 1000, 100, 200),
            OfficeShapedGlyph.CreatePositionedUnicodeContinuation(1, 0, 0, 0, 200, -100)
        }, OfficeTextDirection.LeftToRight);
        prepared = new OfficePreparedTextShaping(program, "A", result);
        drawing.AddPositionedText("A", 10D, 20D, 80D, 40D,
            new OfficeFontInfo(family, 20D), color, textAdvanceWidth: prepared.Measure(20D));
        drawing.MarkLastTextWithPreparedShaping(prepared);
        return drawing;
    }

    private static OfficeDrawing PreparedCopyReference(OfficeColor color, double offsetX, double offsetY, double scale) {
        // At size20 the existing tracking fixture contributes zero. The positioned baseline is40.
        OfficeShape first = OfficeShape.Rectangle(12D * scale, 14D * scale);
        first.FillColor = color;
        first.StrokeColor = null;
        OfficeShape continuation = OfficeShape.Rectangle(8D * scale, 14D * scale);
        continuation.FillColor = color;
        continuation.StrokeColor = null;
        double width = (offsetX == 0D ? 100D : 120D) * scale;
        double height = (offsetY == 0D ? 80D : 100D) * scale;
        return new OfficeDrawing(width, height)
            .AddShape(first, (12D + offsetX) * scale, (22D + offsetY) * scale)
            .AddShape(continuation, (34D + offsetX) * scale, (28D + offsetY) * scale);
    }

    private sealed class RetainedRunProvider : IOfficeTextShapingProvider {
        private readonly OfficeTextShapingResult _result;
        internal RetainedRunProvider(OfficeTextShapingResult result) => _result = result;
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => _result;
    }
}
