using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingContextualKerningTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LogicalContext_KeepsItsBacktrackAndLookaheadAcrossIgnoredMarks(bool extension) {
        var font = Load(ManagedTextShapingTestAssets.CreateContextualPositioningFont(extension));
        var tokens = new[] { Token(1, 'A'), Token(4, 0x0301), Token(2, 'B'), Token(3, 'C') };

        Assert.True(font.TryPositionGlyphRun(tokens, "en", CancellationToken.None, out var positions));

        Assert.Equal(7, positions[2].XPlacement);
        Assert.Equal(-105, positions[2].XAdvance);
        Assert.Equal(20, positions[2].YPlacement);
        Assert.Equal(0, positions[1].XAdvance);
        Assert.Equal(0, positions[1].YPlacement);
        Assert.Equal(0, positions[3].XAdvance);
        // Reversing tokens before positioning cannot satisfy the logical context.
        Assert.True(font.TryPositionGlyphRun(new[] { Token(3, 'C'), Token(2, 'B'), Token(1, 'A') },
            "en", CancellationToken.None, out var reversed));
        Assert.Equal(0, reversed[1].XAdvance);
    }

    [Fact]
    public void LogicalContext_DoesNotConsumeLookaheadOrChangeTheOrdinaryPairRoute() {
        var font = Load(ManagedTextShapingTestAssets.CreateContextualPositioningFont());
        var noLookahead = new[] { Token(1, 'A'), Token(2, 'B') };

        Assert.True(font.TryPositionGlyphRun(noLookahead, "en", CancellationToken.None, out var positions));
        Assert.Equal(0, positions[1].XAdvance);
        OfficeOpenTypeGlyphPositioning[] ordinary = font.PositionGlyphRun(new[] { 1, 2, 3 }, new[] { (int)'A', 'B', 'C' });
        Assert.Equal(-25, ordinary[1].XAdvance);
        Assert.Equal(7, ordinary[1].XPlacement);
    }

    [Fact]
    public void NestedPair_AppliesBothValueRecordsWithinTheMatchedInput() {
        var font = Load(ManagedTextShapingTestAssets.CreateContextualPositioningFont(nestedPair: true));
        var tokens = new[] { Token(1, 'A'), Token(2, 'B'), Token(4, 0x0301), Token(3, 'C') };

        Assert.True(font.TryPositionGlyphRun(tokens, null, CancellationToken.None, out var positions));

        Assert.Equal(-25, positions[1].XAdvance);
        Assert.Equal(7, positions[1].XPlacement);
        Assert.Equal(30, positions[3].YPlacement);
        Assert.Equal(0, positions[2].YPlacement);
    }

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    public void UnsupportedTruncatedOrRecursiveContexts_DeclineTheAdvancedRoute(bool recursive, bool unsupported, bool truncated) {
        var font = Load(ManagedTextShapingTestAssets.CreateContextualPositioningFont(recursive: recursive,
            unsupportedFormat: unsupported, invalidCoverageOffset: truncated));
        Assert.False(font.TryPositionGlyphRun(new[] { Token(1, 'A'), Token(2, 'B'), Token(3, 'C') },
            null, CancellationToken.None, out var positions));
        Assert.Empty(positions);
    }

    [Fact]
    public void LogicalPositioning_ObservesCancellationBeforeInspectingTheFont() {
        var font = Load(ManagedTextShapingTestAssets.CreateContextualPositioningFont());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => font.TryPositionGlyphRun(new[] { Token(1, 'A') },
            null, cancellation.Token, out _));
    }

    private static OfficeTrueTypeFont Load(byte[] bytes) => Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(bytes));
    private static OfficeOpenTypeSubstitution.GlyphToken Token(int glyph, int scalar) =>
        new(glyph, char.ConvertFromUtf32(scalar), 0, scalar);
}
