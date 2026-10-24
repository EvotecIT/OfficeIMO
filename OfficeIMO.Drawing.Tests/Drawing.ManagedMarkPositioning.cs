using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ManagedMarkPositioningTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedAnchors_RetainBaseAndFilteredMarkStackParents(bool extension) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(extension);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301),
            Token(3, "\u0302", 2, 0x0302), Token(4, "\u0303", 3, 0x0303) };

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, "en", OfficeTextFeatureSettings.Default,
            CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));

        Assert.False(positions[0].IsMark);
        Assert.Equal(-1, positions[0].ParentGlyphIndex);
        AssertAttachment(positions[1], 0, 190, 680);
        AssertAttachment(positions[2], 0, 190, 680);
        // The intervening mark is outside the font's mark-filtering set.
        AssertAttachment(positions[3], 1, 39, 78);
    }

    [Theory]
    [InlineData(1, 90, 580)]
    [InlineData(3, 290, 480)]
    public void LigatureAnchors_UseTheMarkSourceComponent(int markIndex, int x, int y) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont();
        object identity = new();
        var ligature = new OfficeOpenTypeSubstitution.GlyphToken(5, "A\u0301A", 0, 'A',
            ligatureComponentTextIndexes: new[] { 0, 2 }, ligatureIdentity: identity);
        var mark = new OfficeOpenTypeSubstitution.GlyphToken(2, "\u0301", markIndex, 0x0301,
            ligatureIdentity: markIndex < 3 ? identity : null, ligatureComponent: markIndex < 3 ? 1 : 0);
        var tokens = new[] { ligature, mark };

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));

        AssertAttachment(positions[1], 0, x, y);
    }

    [Theory]
    [InlineData(1, 90, 580)]
    [InlineData(3, 490, 380)]
    public void PrecomposedLigatureParts_RetainInternalMarkComponentAndFinalExternalAnchor(int markIndex, int x, int y) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(extraLigatureComponent: true);
        object identity = new();
        var ligature = new OfficeOpenTypeSubstitution.GlyphToken(5, "A\u0301A", 0, 'A',
            ligatureComponentTextIndexes: new[] { 0, 2 }, ligatureIdentity: identity);
        var mark = new OfficeOpenTypeSubstitution.GlyphToken(2, "\u0301", markIndex, 0x0301,
            ligatureIdentity: markIndex < 3 ? identity : null, ligatureComponent: markIndex < 3 ? 1 : 0);
        var tokens = new[] { ligature, mark };

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));

        AssertAttachment(positions[1], 0, x, y);
    }

    [Theory]
    [InlineData(0, 0, 490, 380)]
    [InlineData(2, 0, 290, 480)]
    [InlineData(1, 1, 39, 78)]
    public void GraphemeOwnership_DoesNotMergeDistinctLigatureMarkAssociations(int component, int parent, int x, int y) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(extraLigatureComponent: true);
        object identity = new();
        var ligature = new OfficeOpenTypeSubstitution.GlyphToken(5, "A\u0301A\u0303", 0, 'A',
            ligatureComponentTextIndexes: new[] { 0, 2 }, ligatureIdentity: identity);
        var first = new OfficeOpenTypeSubstitution.GlyphToken(2, string.Empty, 0, 0x0301,
            isUnicodeContinuation: true, clusterStart: 0, ligatureIdentity: identity, ligatureComponent: 1);
        var second = new OfficeOpenTypeSubstitution.GlyphToken(4, string.Empty, 2, 0x0303,
            isUnicodeContinuation: true, clusterStart: 0, ligatureIdentity: component == 0 ? null : identity,
            ligatureComponent: component);

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, new[] { ligature, first, second }, null,
            OfficeTextFeatureSettings.Default, CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));

        AssertAttachment(positions[1], 0, 90, 580);
        AssertAttachment(positions[2], parent, x, y);
    }

    [Theory]
    [InlineData(2, false)]
    [InlineData(4, true)]
    public void ParentCoverage_RemainsAuthoritativeWhenLookupFiltersBaseClasses(int flags, bool classifyBaseAsLigature) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(baseLookupFlags: (ushort)flags,
            classifyBaseAsLigature: classifyBaseAsLigature);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301) };

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));

        AssertAttachment(positions[1], 0, 190, 680);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SupportedMissingParentAnchors_RetainTheFontsNominalNoOp(bool missingParentCoverage) {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(missingBaseAnchor: !missingParentCoverage,
            missingParentCoverage: missingParentCoverage);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301) };

        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out OfficeOpenTypeMarkAttachment[] nominal));
        Assert.True(nominal[1].IsMark);
        Assert.True(nominal[1].IsFontDirectedNoOp);
        Assert.Equal(-1, nominal[1].ParentGlyphIndex);
        Assert.Equal(0, nominal[1].DeltaX);
        Assert.Equal(0, nominal[1].DeltaY);
        Assert.True(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null,
            OfficeTextFeatureSettings.Default.With("mark", 0), CancellationToken.None, out OfficeOpenTypeMarkAttachment[] positions));
        Assert.Equal(-1, positions[1].ParentGlyphIndex);
        Assert.True(positions[1].IsMark);
        Assert.False(positions[1].IsFontDirectedNoOp);
    }

    [Fact]
    public void MarksWithoutInterpretedFontCoverage_StillDecline() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 0x0301);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301) };

        Assert.False(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out _));
    }

    [Fact]
    public void UnsupportedDeviceAnchors_DeclineRatherThanClaimStaticCoordinates() {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(unsupportedAnchorDevice: true);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301) };

        Assert.False(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out _));
    }

    [Fact]
    public void UnsupportedDeviceAnchors_CannotBeAdmittedThroughAMissingParent() {
        byte[] font = ManagedTextShapingTestAssets.CreateMarkPositioningFont(unsupportedAnchorDevice: true,
            missingParentCoverage: true);
        var tokens = new[] { Token(1, "A", 0, 'A'), Token(2, "\u0301", 1, 0x0301) };

        Assert.False(OfficeOpenTypeMarkPositioning.TryGetAttachments(font, tokens, null, OfficeTextFeatureSettings.Default,
            CancellationToken.None, out _));
    }

    [Fact]
    public void MarkPositioning_HonorsCancellationBeforeFontInspection() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeOpenTypeMarkPositioning.TryGetAttachments(
            ManagedTextShapingTestAssets.CreateMarkPositioningFont(), new[] { Token(1, "A", 0, 'A') }, null,
            OfficeTextFeatureSettings.Default, cancellation.Token, out _));
    }

    private static OfficeOpenTypeSubstitution.GlyphToken Token(int glyph, string text, int index, int scalar) =>
        new(glyph, text, index, scalar);

    private static void AssertAttachment(OfficeOpenTypeMarkAttachment actual, int parent, int x, int y) {
        Assert.True(actual.IsMark);
        Assert.Equal(parent, actual.ParentGlyphIndex);
        Assert.Equal(x, actual.DeltaX);
        Assert.Equal(y, actual.DeltaY);
    }
}
