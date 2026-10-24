using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Threading;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingManagedResolvedDirectionTests {
    [Fact]
    public void ResolvedRtlLatinSliceReversesWithoutChangingThePublicParagraphContract() {
        byte[] data = ManagedTextShapingTestAssets.CreateControlAdmissionFont(blankSpace: true);
        var paragraphRequest = new OfficeTextShapingRequest("abc", "Resolved", data, false, 1000,
            OfficeTextDirection.RightToLeft, "en", OfficeTextFeatureSettings.Default.With("kern", 0));
        OfficeTextShapingResult paragraph = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(paragraphRequest));
        OfficeTextShapingResult resolved = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(ResolvedRequest("abc", data, OfficeTextDirection.RightToLeft)));

        Assert.Equal("abc", string.Concat(paragraph.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal("cba", string.Concat(resolved.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(new[] { 2, 1, 0 }, resolved.Glyphs.Select(glyph => glyph.TextIndex));
        Assert.Equal("abc", string.Concat(resolved.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
        Assert.Equal(OfficeTextDirection.RightToLeft, resolved.Direction);
    }

    [Theory]
    [InlineData("12")]
    [InlineData("\u0661\u0662")]
    public void ResolvedRtlNumericSliceUsesItsAuthoredOddLevel(string source) {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(source.Select(character => (int)character).ToArray());
        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(ResolvedRequest(source, data, OfficeTextDirection.RightToLeft)));

        Assert.Equal(new[] { 1, 0 }, result.Glyphs.Select(glyph => glyph.TextIndex));
        Assert.Equal(source, string.Concat(result.Glyphs.OrderBy(glyph => glyph.TextIndex).Select(glyph => glyph.UnicodeText)));
    }

    [Fact]
    public void ResolvedLtrHebrewSliceKeepsItsAuthoredLogicalWriteOrder() {
        const string source = "\u05D0\u05D1\u05D2";
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(source.Select(character => (int)character).ToArray());
        OfficeTextShapingResult result = Assert.IsType<OfficeTextShapingResult>(
            OfficeManagedTextShapingProvider.Instance.ShapeText(ResolvedRequest(source, data, OfficeTextDirection.LeftToRight)));

        Assert.Equal(source, string.Concat(result.Glyphs.Select(glyph => glyph.UnicodeText)));
        Assert.Equal(new[] { 0, 1, 2 }, result.Glyphs.Select(glyph => glyph.TextIndex));
        Assert.Equal(OfficeTextDirection.LeftToRight, result.Direction);
    }

    [Fact]
    public void ResolvedOddLevelLatinStillDeclinesFontDirectedSubstitution() {
        byte[] data = ManagedTextShapingTestAssets.CreateControlAdmissionFont(blankSpace: true, withSubstitution: true);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(ResolvedRequest("abc", data, OfficeTextDirection.RightToLeft)));
    }

    [Fact]
    public void ForcedResolvedSliceCannotFallThroughToTheDiagnosedParagraphFallback() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('a', 'b', 0x0301);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeDiagnosedFallback(
            ResolvedRequest("a\u0301b", data, OfficeTextDirection.RightToLeft)));
    }

    [Theory]
    [InlineData(OfficeTextDirection.Auto)]
    [InlineData(OfficeTextDirection.TopToBottom)]
    public void ResolvedSliceRequiresAConcreteHorizontalDirection(OfficeTextDirection direction) {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('a');

        Assert.Throws<ArgumentException>(() => ResolvedRequest("a", data, direction));
    }

    [Theory]
    [InlineData("\u202Eabc\u202C")]
    [InlineData("a\uD800")]
    public void ResolvedSliceDeclinesEmbeddingDirectivesAndInvalidUtf16(string source) {
        byte[] data = ManagedTextShapingTestAssets.CreateControlAdmissionFont(blankSpace: true);

        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(ResolvedRequest(source, data, OfficeTextDirection.RightToLeft)));
    }

    [Fact]
    public void ResolvedSliceRetainsTheNormalizationInputBoundAndCancellation() {
        byte[] data = ManagedTextShapingTestAssets.CreateFont('a');
        Assert.Null(OfficeManagedTextShapingProvider.Instance.ShapeText(
            ResolvedRequest(new string('a', 1_000_001), data, OfficeTextDirection.RightToLeft)));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => OfficeManagedTextShapingProvider.Instance.ShapeText(
            ResolvedRequest("a", data, OfficeTextDirection.RightToLeft, cancellation.Token)));
    }

    private static OfficeTextShapingRequest ResolvedRequest(string source, byte[] data, OfficeTextDirection direction,
        CancellationToken cancellationToken = default) => new(source, "Resolved", data, false, 1000,
            direction, "en", cancellationToken, null, null, cloneFontData: false,
            featureSettings: OfficeTextFeatureSettings.Default.With("kern", 0), resolvedDirectionalRun: true);
}
