using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingManagedTypographyCompositionTests {
    [Fact]
    public void AlternatingScriptsShareOneSubstitutionWorkBudget() {
        byte[] data = ManagedTextShapingTestAssets.CreateAlternatingScriptBudgetFont();
        string ordinary = string.Concat(Enumerable.Repeat("بA", 4));
        string excessive = string.Concat(Enumerable.Repeat("بA", 1024));

        Assert.IsType<OfficeTextShapingResult>(Shape(data, ordinary));
        // Every individual segment is within the bound; the complete public
        // request must not obtain a fresh work allowance for every script switch.
        Assert.Null(Shape(data, excessive));
    }

    [Fact]
    public void SharedScriptContextHonorsCancellationBeforeSourceAnalysis() {
        byte[] data = ManagedTextShapingTestAssets.CreateAlternatingScriptBudgetFont();
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Cancelled composition", data).Faces).Program;
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => OfficeManagedTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest("بA", "Cancelled composition", data, false, font.UnitsPerEm,
                OfficeTextDirection.LeftToRight, "ar", cancellationToken: cancellation.Token)));
    }

    [Fact]
    public void MarkAndStackFeaturesUseCombinedLookupListOrder() {
        byte[] data = ManagedTextShapingTestAssets.CreateMarkCompositionFont(stackBeforeBase: true);
        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(Shape(data, "A\u0301\u0303"));

        Assert.Equal(new[] { 1, 2, 4 }, result.Glyphs.Select(glyph => glyph.GlyphId));
        // Lookup 0 first stacks the second mark. Lookup 1 subsequently anchors both
        // marks to A, so its later attachment must replace the earlier mark parent.
        Assert.All(result.Glyphs.Skip(1), glyph => {
            Assert.Equal(-310, glyph.OffsetX); Assert.Equal(680, glyph.OffsetY);
            Assert.Equal(0, glyph.AdvanceWidth);
        });
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(4, false)]
    [InlineData(8, false)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    [InlineData(4, true)]
    [InlineData(8, true)]
    public void MarkPositioningCannotSilentlyDiscardSingleOrPairAdjustments(int field, bool pair) {
        byte[] data = ManagedTextShapingTestAssets.CreateMarkCompositionFont(field, pair);

        Assert.Null(Shape(data, "A\u0301"));
        Assert.IsType<OfficeTextShapingResult>(Shape(data, "A\u0301", OfficeTextFeatureSettings.Default.With("kern", 0)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SelectedUnsupportedPositioningFeaturesDecline(bool required) {
        byte[] data = ManagedTextShapingTestAssets.CreateMarkCompositionFont(1, extraFeature: "dist", requiredExtraFeature: required);

        Assert.Null(Shape(data, "A\u0301", required ? OfficeTextFeatureSettings.Default
            : OfficeTextFeatureSettings.Default.With("dist", 1)));
    }

    private static OfficeTextShapingResult? Shape(byte[] data, string text, OfficeTextFeatureSettings? features = null) {
        IOfficeFontProgram font = Assert.Single(new OfficeFontFaceCollection().Add("Positioning composition", data).Faces).Program;
        return OfficeManagedTextShapingProvider.Instance.ShapeText(new OfficeTextShapingRequest(text,
            "Positioning composition", data, false, font.UnitsPerEm, OfficeTextDirection.LeftToRight, "en", featureSettings: features));
    }
}
