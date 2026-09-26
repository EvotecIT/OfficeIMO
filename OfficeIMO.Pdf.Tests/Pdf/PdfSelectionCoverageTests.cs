using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSelectionCoverageTests {
    [Fact]
    public void CoverageUsesTrueQuadAreaInsteadOfItsBoundingRectangle() {
        PdfSelectionQuad target = Diamond(5, 5, 5);
        var corners = new[] { Rectangle(0, 0, 4, 4), Rectangle(6, 0, 10, 4),
            Rectangle(0, 6, 4, 10), Rectangle(6, 6, 10, 10) };
        Assert.False(Covers(target, corners, 0.5D));
        Assert.True(Covers(target, corners, 0.35D));
    }

    [Fact]
    public void CoverageUnionsClippedPolygonsWithoutCountingDuplicateNativeSpansTwice() {
        PdfSelectionQuad target = Diamond(5, 5, 5);
        PdfSelectionQuad left = Rectangle(0, 0, 5, 10), right = Rectangle(5, 0, 10, 10);
        Assert.False(Covers(target, new[] { left, left, left }, 0.6D));
        Assert.True(Covers(target, new[] { left, right }, 0.99D));
    }

    [Fact]
    public void CoverageAccountsForCrossingEdgesBetweenNativeQuads() {
        // Each diamond has area 18; their overlapping diamond has area 2, giving union area 34.
        PdfSelectionQuad target = Rectangle(0, 0, 10, 10);
        var native = new[] { Diamond(3, 5, 3), Diamond(7, 5, 3) };
        Assert.True(Covers(target, native, 0.33D));
        Assert.False(Covers(target, native, 0.35D));
    }

    [Fact]
    public void EmptyOrDisjointNativeTextNeverRejectsWordsEvenAtZeroThreshold() {
        PdfSelectionQuad target = Rectangle(0, 0, 10, 10);
        Assert.False(Covers(target, Array.Empty<PdfSelectionQuad>(), 0D));
        Assert.False(Covers(target, new[] { Rectangle(20, 20, 30, 30) }, 0D));
        Assert.True(Covers(target, new[] { Rectangle(0, 0, 1, 1) }, 0D));
    }

    [Fact]
    public void GeometryWorkHonorsCallerBudgetAndCancellation() {
        PdfSelectionQuad target = Diamond(5, 5, 5);
        Assert.Throws<InvalidOperationException>(() => PdfSelectionCoverage.Covers(target,
            new[] { Rectangle(0, 0, 5, 10) }, 0.6D, work => {
                if (work > 1) throw new InvalidOperationException("Geometry budget exhausted.");
            }, default));
        Assert.ThrowsAny<OperationCanceledException>(() => PdfSelectionCoverage.Covers(target,
            new[] { target }, 0.5D, _ => { }, new CancellationToken(true)));
    }

    [Fact]
    public void CoverageBoundsRetainedIntersectionsBeforeUnionAllocation() {
        PdfSelectionQuad target = Rectangle(0, 0, 10, 10);
        PdfSelectionQuad fragment = Rectangle(0, 0, 1, 1);

        PdfReadLimitException error = Assert.Throws<PdfReadLimitException>(() => PdfSelectionCoverage.Covers(
            target, new[] { fragment, fragment, fragment }, 0.9D, _ => { }, default,
            maximumRetainedIntersections: 2));

        Assert.Equal(PdfReadLimitKind.OcrArtifacts, error.Kind);
        Assert.Equal(3, error.Actual);
    }

    [Fact]
    public void CoveringIntersectionDoesNotNeedAnotherRetainedSlot() {
        PdfSelectionQuad rectangle = Rectangle(0, 0, 10, 10);
        Assert.True(PdfSelectionCoverage.Covers(rectangle,
            new[] { Rectangle(0, 0, 1, 1), rectangle }, 0.9D, _ => { }, default,
            maximumRetainedIntersections: 1));

        PdfSelectionQuad diamond = Diamond(5, 5, 5);
        Assert.True(PdfSelectionCoverage.Covers(diamond,
            new[] { Rectangle(0, 0, 5, 10), diamond }, 0.9D, _ => { }, default,
            maximumRetainedIntersections: 1));
    }

    [Fact]
    public void CoverageDoesNotChargeEmptyPolygonIntersections() {
        PdfSelectionQuad target = Diamond(5, 5, 5);
        PdfSelectionQuad left = Rectangle(0, 0, 5, 10);
        PdfSelectionQuad outsideCorner = Rectangle(0, 0, 1, 1);

        Assert.False(PdfSelectionCoverage.Covers(target, new[] { left, outsideCorner },
            0.9D, _ => { }, default, maximumRetainedIntersections: 1));
    }

    private static bool Covers(PdfSelectionQuad target, IReadOnlyList<PdfSelectionQuad> native, double threshold) =>
        PdfSelectionCoverage.Covers(target, native, threshold, _ => { }, default);

    private static PdfSelectionQuad Rectangle(double left, double top, double right, double bottom) => new PdfSelectionQuad(
        new PdfSelectionPoint(left, top), new PdfSelectionPoint(right, top),
        new PdfSelectionPoint(right, bottom), new PdfSelectionPoint(left, bottom));

    private static PdfSelectionQuad Diamond(double x, double y, double radius) => new PdfSelectionQuad(
        new PdfSelectionPoint(x, y - radius), new PdfSelectionPoint(x + radius, y),
        new PdfSelectionPoint(x, y + radius), new PdfSelectionPoint(x - radius, y));
}
