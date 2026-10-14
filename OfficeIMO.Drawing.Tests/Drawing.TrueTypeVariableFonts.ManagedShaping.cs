using System.Collections.Generic;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed partial class DrawingTrueTypeVariableFontTests {
    [Theory]
    [InlineData(400F, 60)]
    [InlineData(900F, 10)]
    public void ManagedPairPositioningUsesSelectedVariableFontCoordinates(
        float weight,
        int expectedAdjustment) {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        var axes = new Dictionary<string, float> { ["wght"] = weight };
        OfficeFontFace selected = Load(data, axes);
        const string pair = "\"V";
        var request = new OfficeTextShapingRequest(
            pair,
            "Roboto Flex",
            data,
            isOpenTypeCff: false,
            unitsPerEm: selected.Program.UnitsPerEm,
            direction: OfficeTextDirection.LeftToRight,
            language: "en",
            cancellationToken: CancellationToken.None,
            fontCollectionIndex: null,
            variationCoordinates: axes,
            featureSettings: OfficeTextFeatureSettings.Default.With("kern", 1));

        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(request));

        AssertVariablePairMeasurement(selected.Program, pair, result, expectedAdjustment);
    }

    [Fact]
    public void DefaultLatinPairPositioningDoesNotReuseAnotherVariableFontInstance() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        const string pair = "\"V";
        // One byte snapshot deliberately exercises the default Latin font cache.
        // Repeating each instance after the other catches cache-order contamination.
        foreach (float weight in new[] { 400F, 900F, 400F, 900F }) {
            var axes = new Dictionary<string, float> { ["wght"] = weight };
            OfficeFontFace selected = Load(data, axes);
            var request = new OfficeTextShapingRequest(
                pair,
                "Roboto Flex",
                data,
                isOpenTypeCff: false,
                unitsPerEm: selected.Program.UnitsPerEm,
                direction: OfficeTextDirection.LeftToRight,
                language: "en",
                cancellationToken: CancellationToken.None,
                fontCollectionIndex: null,
                variationCoordinates: axes,
                cloneFontData: false,
                featureSettings: OfficeTextFeatureSettings.Default.With("kern", 1),
                applyDefaultLatinLigatures: true);

            OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
                OfficeManagedTextShapingProvider.Instance.ShapeText(request));

            AssertVariablePairMeasurement(selected.Program, pair, result,
                weight == 900F ? 10 : 60);
        }
    }

    private static void AssertVariablePairMeasurement(
        IOfficeFontProgram selected,
        string pair,
        OfficeTextShapingResult result,
        int expectedAdjustment) {
        double fontSize = selected.UnitsPerEm;
        double nominalAdvance = selected.Measure(pair[0].ToString(), fontSize)
            + selected.Measure(pair[1].ToString(), fontSize);

        Assert.Equal(pair, string.Concat(result.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(2, result.Glyphs.Count);
        // Existing selected-program proof gives +60 at default and +10 at900.
        // The public shaped measurement must retain the same GDEF variation delta.
        Assert.Equal(expectedAdjustment, PairAdjustment(selected, pair), 6);
        Assert.Equal(nominalAdvance + expectedAdjustment,
            selected.MeasureShapedText(pair, result, fontSize), 6);
    }
}
