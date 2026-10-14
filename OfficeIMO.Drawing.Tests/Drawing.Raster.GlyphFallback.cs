using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterGlyphFallbackTests {
    private const double SymbolSize = 28D;
    private const string BulletSymbols = "•◆■□❖➢✔";

    [Fact]
    public void TrimmingReselectsFallbackWhenEllipsisGlyphsAreUncovered() {
        var actual = new OfficeRasterImage(120, 50);
        var expected = new OfficeRasterImage(120, 50);
        var canvas = CreateSparseCanvas(actual);
        const double size = 20D;
        double width = canvas.MeasureText("A...", size) + 6.5D;
        canvas.DrawText("AAAAAAAAAAAA", 0D, 0D, width, 40D, OfficeColor.Black, size);
        CreateSparseCanvas(expected).DrawText("A...", 0D, 0D, width, 40D, OfficeColor.Black, size);
        Assert.Equal(expected.GetPixels(), actual.GetPixels());
        Assert.True(InkBounds(actual).Pixels > 0);
    }

    [Theory]
    [InlineData("\U0001F469\u200D\U0001F4BB")]
    [InlineData("\u0915\u093F")]
    public void UncoveredComplexTextReportsOneIncompleteFallback(string text) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(120, 70),
            font: OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateFont('A')),
            fonts: null, diagnosticSink: diagnostics, diagnosticSource: "uncovered shaping test");
        canvas.MeasureText(text, 18D);
        canvas.DrawTextLine(text, 20D, 20D, 18D, OfficeColor.Black);
        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.TextShapingFallback, diagnostic.Code);
        Assert.Contains("cannot provide complete", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind);
        Assert.Equal("uncovered shaping test", diagnostic.Source);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("OfficeIMO Missing Glyph Fixture")]
    public void UncoveredFontUsesVisibleDistinctBulletFallbackAndItsMetrics(string? family) {
        var signatures = new HashSet<string>(StringComparer.Ordinal);
        foreach (char symbol in BulletSymbols + "?") {
            var image = new OfficeRasterImage(100, 70);
            var canvas = CreateSparseCanvas(image);
            Assert.Equal(20D, canvas.MeasureText(symbol.ToString(), SymbolSize, family), 10);
            canvas.DrawTextLine(symbol.ToString(), 20D, 20D, SymbolSize, OfficeColor.Black,
                alignment: OfficeTextAlignment.Left, fontFamily: family);
            var bounds = InkBounds(image);
            Assert.True(bounds.Pixels > 0, $"Missing ink for U+{(int)symbol:X4}.");
            Assert.InRange(bounds.Right, 20, 40);
            Assert.True(signatures.Add(Convert.ToBase64String(image.GetPixels())),
                $"U+{(int)symbol:X4} lost its marker identity.");
        }
    }

    [Fact]
    public void BulletFallbackKeepsFilledAndHollowSquareSemantics() {
        var filled = new OfficeRasterImage(100, 70);
        var hollow = new OfficeRasterImage(100, 70);
        CreateSparseCanvas(filled).DrawTextLine("■", 20D, 20D, SymbolSize, OfficeColor.Black,
            alignment: OfficeTextAlignment.Left);
        CreateSparseCanvas(hollow).DrawTextLine("□", 20D, 20D, SymbolSize, OfficeColor.Black,
            alignment: OfficeTextAlignment.Left);
        Assert.True(filled.GetPixel(30, 34).A > 0);
        Assert.Equal(0, hollow.GetPixel(30, 34).A);
        Assert.True(InkBounds(hollow).Pixels > 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UncoveredFontPaintsBulletThroughRotatedAndAffineRoutes(bool affine) {
        foreach (char symbol in BulletSymbols) {
            var actual = new OfficeRasterImage(100, 100);
            var canvas = CreateSparseCanvas(actual);
            if (affine) {
                canvas.DrawTextLineTransformed(symbol.ToString(), 20D, 20D, SymbolSize, OfficeColor.Black,
                    new OfficeTransform(0D, 1D, -1D, 0D, 80D, 0D), alignment: OfficeTextAlignment.Left);
            } else {
                canvas.DrawTextLine(symbol.ToString(), 20D, 20D, SymbolSize, OfficeColor.Black,
                    alignment: OfficeTextAlignment.Left, rotationDegrees: 90D,
                    rotationCenterX: 40D, rotationCenterY: 40D);
            }
            var bounds = InkBounds(actual);
            Assert.True(bounds.Pixels > 0, $"Missing transformed ink for U+{(int)symbol:X4}.");
            Assert.InRange(bounds.Left, 30, 65);
            Assert.InRange(bounds.Top, 20, 40);
            Assert.InRange(bounds.Bottom, 20, 40);
        }
    }

    [Theory]
    [InlineData(28D, 40D)]
    [InlineData(4D, 12D)]
    [InlineData(.1D, 12D)]
    public void FallbackBoundsContainStyledInkBeyondAThinPositionedFrame(double size, double advance) {
        var image = new OfficeRasterImage(160, 120);
        var canvas = CreateSparseCanvas(image);
        OfficeFontStyle style = OfficeFontStyle.Bold | OfficeFontStyle.Italic;
        var font = new OfficeFontInfo(string.Empty, size, style);
        var bounds = canvas.MeasurePositionedTextBounds("❖", 10D, 10D, 1D, 1D, size,
            font, advance, OfficeTextAlignment.Left, OfficeTextFeatureSettings.Default,
            "normal", size, OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None);
        canvas.DrawPositionedText("❖", 10D, 10D, 1D, 1D, OfficeColor.Black, size,
            OfficeTextAlignment.Left, style, null, advance,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, baselineFontSize: size);
        var ink = InkBounds(image);
        Assert.True(bounds.HasInk && ink.Pixels > 0);
        Assert.True(ink.Left >= Math.Floor(bounds.Left) && ink.Right <= Math.Ceiling(bounds.Right));
        Assert.True(ink.Top >= Math.Floor(bounds.Top) && ink.Bottom <= Math.Ceiling(bounds.Bottom));
        var paint = canvas.MeasureTextPaintBounds("❖", size, null, style);
        var natural = new OfficeRasterImage(160, 120);
        var naturalCanvas = CreateSparseCanvas(natural);
        naturalCanvas.DrawPositionedText("❖", 10D, 10D, 1D, 1D, OfficeColor.Black, size,
            OfficeTextAlignment.Left, style, null, naturalCanvas.MeasureText("❖", size),
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, baselineFontSize: size * .84D);
        var naturalInk = InkBounds(natural);
        double baseline = 10D + size * .84D;
        Assert.True(naturalInk.Top >= Math.Floor(paint.Top + baseline));
        Assert.True(naturalInk.Bottom <= Math.Ceiling(paint.Bottom + baseline));
    }

    [Fact]
    public void UncoveredRectangleTextRetainsRequestedSizeAndAlignment() {
        var image = new OfficeRasterImage(160, 120);
        var canvas = CreateSparseCanvas(image);
        canvas.DrawText("◆", 10D, 10D, 120D, 80D, OfficeColor.Black, SymbolSize,
            OfficeTextAlignment.Right);
        var bounds = InkBounds(image);
        Assert.True(bounds.Pixels > 0);
        Assert.InRange(bounds.Right, 122, 129);
        Assert.InRange(bounds.Bottom - bounds.Top, 1, 28);
    }

    [Theory]
    [InlineData(28D, 40D)]
    [InlineData(4D, 12D)]
    public void UncoveredPositionedTextRetainsBaselineAndAuthoredAdvance(double size, double advance) {
        var image = new OfficeRasterImage(160, 120);
        var canvas = CreateSparseCanvas(image);
        canvas.DrawPositionedText("◆", 10D, 10D, 120D, 80D, OfficeColor.Black, size,
            OfficeTextAlignment.Right, OfficeFontStyle.Regular, null, advance,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, baselineFontSize: size);
        var bounds = InkBounds(image);
        Assert.True(bounds.Pixels > 0);
        Assert.InRange(bounds.Left, (int)(130D - advance), 130);
        Assert.InRange(bounds.Right, (int)(130D - advance / 3D), 132);
        Assert.InRange(bounds.Bottom, 11, (int)Math.Ceiling(10D + size * 1.4D + 2D));
    }

    [Fact]
    public void CoveredGlyphRetainsItsFontMetricsAndOutline() {
        OfficeTrueTypeFont font = OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateFont('A'))!;
        var image = new OfficeRasterImage(100, 70);
        var canvas = new OfficeRasterCanvas(image, font: font);
        Assert.Equal(font.Measure("A", SymbolSize), canvas.MeasureText("A", SymbolSize));
        Assert.NotEqual(20D, canvas.MeasureText("A", SymbolSize));
        canvas.DrawTextLine("A", 20D, 20D, SymbolSize, OfficeColor.Black, alignment: OfficeTextAlignment.Left);
        Assert.True(InkBounds(image).Pixels > 0);
    }

    private static OfficeRasterCanvas CreateSparseCanvas(OfficeRasterImage image) =>
        new(image, font: OfficeTrueTypeFont.TryLoad(ManagedTextShapingTestAssets.CreateFont('A')));

    private static (int Left, int Top, int Right, int Bottom, int Pixels) InkBounds(OfficeRasterImage image) {
        int left = image.Width, top = image.Height, right = -1, bottom = -1, pixels = 0;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                if (image.GetPixel(x, y).A == 0) continue;
                left = Math.Min(left, x); top = Math.Min(top, y);
                right = Math.Max(right, x); bottom = Math.Max(bottom, y); pixels++;
            }
        }
        return (left, top, right, bottom, pixels);
    }
}
