using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontOpticalSizeTests {
    private const string Text = "Variable OfficeIMO";

    [Theory]
    [InlineData(0.49D, 12F)]
    [InlineData(0.51D, 12.0000152587890625F)]
    public void AuthoredOpticalSizeKeepsDoublePrecisionUntilFixedSelection(double fractionalStep, float expectedSize) {
        double authoredSize = 12D + fractionalStep / 65536D;
        OfficeFontFaceCollection automatic = Fonts(null);
        OfficeFontFace selected = Resolve(automatic, authoredSize);
        OfficeFontFace expected = Assert.Single(Fonts(expectedSize).Faces);

        Assert.Equal(expected.Program.Fingerprint, selected.Program.Fingerprint);
        Assert.Equal(expected.Program.Measure(Text, authoredSize), selected.Program.Measure(Text, authoredSize));
        Assert.Equal(expected.Program.GetTextContours(Text, 0D, 0D, authoredSize),
            selected.Program.GetTextContours(Text, 0D, 0D, authoredSize));
        Assert.Same(selected, Resolve(automatic, expectedSize));
    }

    [Theory]
    [InlineData(512D)]
    [InlineData(512.0000152587890625D)]
    public void AutomaticOpticalInstancesKeepCanonicalCoordinatesBeyondFloatPrecision(double defaultSize) {
        const double step = 1D / 65536D;
        var fonts = new OfficeFontFaceCollection();
        Assert.True(fonts.TryAdd("Probe", ManagedTextShapingTestAssets.CreateNormalizationFont(
            axis: ("opsz", 512D, defaultSize, 512.125D))));
        OfficeFontFace ResolveProbe(double size) => Resolve(fonts, size, "A");
        OfficeFontFace original = Assert.Single(fonts.Faces);
        OfficeFontFace selected = ResolveProbe(defaultSize + step);
        OfficeFontFace adjacent = ResolveProbe(defaultSize + 2D * step);

        Assert.Same(original, ResolveProbe(defaultSize));
        Assert.NotSame(original, selected);
        Assert.NotSame(selected, adjacent);
        Assert.NotEqual(original.Program.Fingerprint, selected.Program.Fingerprint);
        Assert.NotEqual(selected.Program.Fingerprint, adjacent.Program.Fingerprint);
        Assert.Contains("opsz=" + (defaultSize + step).ToString("R", CultureInfo.InvariantCulture), selected.Program.Fingerprint);
        Assert.Same(selected, ResolveProbe(defaultSize + 1.49D * step));
        Assert.Same(adjacent, ResolveProbe(defaultSize + 1.51D * step));
        Assert.Equal(0.03125D, selected.Program.GetTextContours("A", 0D, 0D, selected.Program.UnitsPerEm)
            .SelectMany(contour => contour).Min(point => point.X));
        Assert.Equal(0.0625D, adjacent.Program.GetTextContours("A", 0D, 0D, adjacent.Program.UnitsPerEm)
            .SelectMany(contour => contour).Min(point => point.X));
        Assert.Equal(502D, selected.Program.Measure("A", selected.Program.UnitsPerEm));
        Assert.Equal(504D, adjacent.Program.Measure("A", adjacent.Program.UnitsPerEm));
    }

    [Fact]
    public void ManagedPositioningRetainsTheSelectedAutomaticOpticalInstance() {
        const string pair = "\"V";
        const double size = 512D + 1D / 65536D;
        byte[] data = ManagedTextShapingTestAssets.WithOpticalSizeRange(
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf")), 512D, 512D, size);
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F }
        };
        Assert.True(fonts.TryAdd("Probe", data));
        IOfficeFontProgram selected = Resolve(fonts, size, pair).Program;
        var features = OfficeTextFeatureSettings.Default.With("kern", 1);
        // Public requests keep their float contract. An out-of-range float clamps to
        // the exact authored maximum, providing an independent endpoint reference.
        var referenceRequest = new OfficeTextShapingRequest(pair, "Probe", data, false, selected.UnitsPerEm,
            OfficeTextDirection.LeftToRight, "en", default, null,
            new Dictionary<string, float> { ["wght"] = 700F, ["opsz"] = 513F }, features);
        OfficeTextShapingResult reference = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(referenceRequest));
        // This existing internal constructor is the real HTML/raster request seam:
        // the selected immutable program accompanies its public float projection.
        var actualRequest = new OfficeTextShapingRequest(pair, "Probe", selected.GetFontDataForShaping(), false,
            selected.UnitsPerEm, OfficeTextDirection.LeftToRight, "en", default, null,
            ((IOfficeVariableFontProgram)selected).VariationCoordinatesForShaping,
            cloneFontData: false, fontProgramCacheKey: selected, featureSettings: features);
        OfficeTextShapingResult actual = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(actualRequest));

        Assert.Equal(512F, actualRequest.VariationCoordinates["opsz"]);
        Assert.Equal(2122D, selected.MeasureShapedText(pair, reference, selected.UnitsPerEm));
        Assert.Equal(selected.MeasureShapedText(pair, reference, selected.UnitsPerEm),
            selected.MeasureShapedText(pair, actual, selected.UnitsPerEm));
    }

    [Theory]
    [InlineData(10D)]
    [InlineData(16D)]
    [InlineData(24D)]
    [InlineData(48D)]
    [InlineData(200D)]
    public void AuthoredSizeMatchesAnExplicitOpticalInstanceAndPreservesWeight(double size) {
        OfficeFontFaceCollection automatic = Fonts(null);
        OfficeFontFaceCollection reference = Fonts((float)size);
        Assert.True(automatic.TryResolveFaceForText(Text, "Probe", OfficeFontStyle.Regular, size, out OfficeFontFace? selected));
        OfficeFontFace expected = Assert.Single(reference.Faces);
        Assert.Equal(expected.Program.Fingerprint, selected!.Program.Fingerprint);
        Assert.False(selected.CanEmbedAsStaticPdfFont);
        Assert.True(automatic.TryMeasureText(Text, size, "Probe", OfficeFontStyle.Regular, out double measured));
        Assert.Equal(expected.Program.Measure(Text, size), measured);
        Assert.Equal(expected.Program.GetTextContours(Text, 0, 0, size), selected.Program.GetTextContours(Text, 0, 0, size));
        Assert.Same(selected, Resolve(automatic, size));
    }

    [Fact]
    public void ExplicitOpticalSelectionWinsOverAuthoredSize() {
        OfficeFontFaceCollection fonts = Fonts(20F);
        OfficeFontFace original = Assert.Single(fonts.Faces);
        Assert.Same(original, Resolve(fonts, 10D));
        Assert.Same(original, Resolve(fonts, 48D));
        Assert.True(fonts.TryMeasureText(Text, 48D, "Probe", OfficeFontStyle.Regular, out double width));
        Assert.Equal(original.Program.Measure(Text, 48D), width);
    }

    [Fact]
    public void SuccessiveRasterMetricScopesRetainTheirOwnOpticalMeasurementsAndInk() {
        OfficeRasterCanvas Canvas() => new(new OfficeRasterImage(1, 1), fonts: Fonts(null));
        OfficeRasterCanvas reused = Canvas();
        double originalWidth = reused.MeasureText(Text, 24D, "Probe", OfficeFontStyle.Regular);
        var originalInk = reused.MeasureTextLineInkBounds(Text, 24D, "Probe", OfficeFontStyle.Regular);
        var originalPaint = reused.MeasureTextPaintBounds("Ágj", 24D, "Probe", OfficeFontStyle.Regular);
        using (reused.PushFontMetricScale(2D)) {
            OfficeRasterCanvas reference = Canvas();
            using (reference.PushFontMetricScale(2D)) {
                double expectedWidth = reference.MeasureText(Text, 24D, "Probe", OfficeFontStyle.Regular);
                var expectedInk = reference.MeasureTextLineInkBounds(Text, 24D, "Probe", OfficeFontStyle.Regular);
                var expectedPaint = reference.MeasureTextPaintBounds("Ágj", 24D, "Probe", OfficeFontStyle.Regular);
                Assert.NotEqual(originalWidth, expectedWidth);
                Assert.NotEqual(originalInk, expectedInk);
                Assert.NotEqual((originalPaint.Top, originalPaint.Bottom), (expectedPaint.Top, expectedPaint.Bottom));
                Assert.Equal(expectedWidth, reused.MeasureText(Text, 24D, "Probe", OfficeFontStyle.Regular));
                Assert.Equal(expectedInk, reused.MeasureTextLineInkBounds(Text, 24D, "Probe", OfficeFontStyle.Regular));
                var actualPaint = reused.MeasureTextPaintBounds("Ágj", 24D, "Probe", OfficeFontStyle.Regular);
                Assert.Equal((expectedPaint.Top, expectedPaint.Bottom), (actualPaint.Top, actualPaint.Bottom));
            }
        }
        Assert.Equal(originalWidth, reused.MeasureText(Text, 24D, "Probe", OfficeFontStyle.Regular));
        Assert.Equal(originalInk, reused.MeasureTextLineInkBounds(Text, 24D, "Probe", OfficeFontStyle.Regular));
        var restoredPaint = reused.MeasureTextPaintBounds("Ágj", 24D, "Probe", OfficeFontStyle.Regular);
        Assert.Equal((originalPaint.Top, originalPaint.Bottom), (restoredPaint.Top, restoredPaint.Bottom));
    }

    [Fact]
    public void ProviderOwnedTrueTypeProgramKeepsItsSelectedInstance() {
        IOfficeFontProgram program = Assert.Single(Fonts(null).Faces).Program;
        var fonts = new OfficeFontFaceCollection { FontProgramProvider = new FixedProvider(program) };
        Assert.True(fonts.TryAdd("Probe", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"))));
        foreach (OfficeFontFaceCollection snapshot in new[] { fonts, fonts.Clone() }) {
            Assert.Same(program, Resolve(snapshot, 48D).Program);
            Assert.True(snapshot.TryMeasureText(Text, 48D, "Probe", OfficeFontStyle.Regular, out double width));
            Assert.Equal(program.Measure(Text, 48D), width);
        }
    }

    [Theory]
    [InlineData(16D, 1D)]
    [InlineData(16D, 2D)]
    [InlineData(24D, 1D)]
    [InlineData(24D, 2D)]
    public void RasterOutputUsesAuthoredOpticalSizeBeforeOutputScale(double size, double scale) {
        OfficeFontFaceCollection automatic = Fonts(null);
        OfficeFontFaceCollection reference = Fonts((float)size);
        byte[] Render(OfficeFontFaceCollection fonts) {
            Assert.True(fonts.TryMeasureText(Text, size, "Probe", OfficeFontStyle.Regular, out double width));
            var drawing = new OfficeDrawing(500D, 100D).AddPositionedText(Text, 10D, 10D, width, 60D,
                new OfficeFontInfo("Probe", size), OfficeColor.Black, textAdvanceWidth: width);
            drawing.Fonts.AddRange(fonts);
            return OfficeDrawingRasterRenderer.Render(drawing, scale, OfficeColor.White).GetPixels();
        }
        Assert.Equal(Render(reference), Render(automatic));
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(-1D)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void SizeAwareFaceResolutionRejectsInvalidSizes(double size) {
        Assert.Throws<ArgumentOutOfRangeException>(() => Fonts(null).TryResolveFaceForText(
            Text, "Probe", OfficeFontStyle.Regular, size, out _));
    }

    private static OfficeFontFace Resolve(OfficeFontFaceCollection fonts, double size, string text = Text) {
        Assert.True(fonts.TryResolveFaceForText(text, "Probe", OfficeFontStyle.Regular, size, out OfficeFontFace? face));
        return face!;
    }

    private static OfficeFontFaceCollection Fonts(float? opticalSize) {
        var axes = new Dictionary<string, float> { ["wght"] = 700F };
        if (opticalSize.HasValue) axes.Add("opsz", opticalSize.Value);
        var fonts = new OfficeFontFaceCollection { FontVariationResolver = _ => axes };
        Assert.True(fonts.TryAdd("Probe", File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"))));
        return fonts;
    }

    private sealed class FixedProvider(IOfficeFontProgram program) : IOfficeFontProgramProvider {
        public OfficeFontProgramLoadResult TryLoad(OfficeFontProgramLoadRequest request) =>
            new(program, request.Data.Length);
    }
}
