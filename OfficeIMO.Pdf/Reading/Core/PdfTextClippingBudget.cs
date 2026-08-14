using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal sealed class PdfTextClippingBudget {
    private int _pathCount;
    private long _intersectionWork;

    internal void ChargePath() {
        PdfPageClipPath.ThrowIfTextClippingPathBudgetExceeded(_pathCount);
        _pathCount++;
    }

    internal PdfPageClipPath ResolveActiveClip(PdfPageClipPath? activeClipPath, PdfPageClipPath clipPath) {
        return PdfPageClipPath.ResolveActiveClip(activeClipPath, clipPath, this);
    }

    internal void ChargeFlattenedPathWork(
        IReadOnlyList<OfficePathCommand> subjectCommands,
        IReadOnlyList<OfficePathCommand> clipCommands) {
        long subjectVertices = PdfPageClipPath.CountFlattenedPathVertices(subjectCommands);
        long clipVertices = PdfPageClipPath.CountFlattenedPathVertices(clipCommands);
        ChargeIntersectionWork(SaturatingAdd(subjectVertices, clipVertices));
    }

    internal void ChargeContourBoundsWork(IReadOnlyList<List<OfficePoint>> contours) {
        ChargeIntersectionWork(CountVertices(contours));
    }

    internal void ChargeLinearIntersectionWork(int pathCommandCount) {
        ChargeIntersectionWork(pathCommandCount);
    }

    internal void ChargeLinearIntersectionWork(long work) {
        ChargeIntersectionWork(work);
    }

    private void ChargeIntersectionWork(long addedWork) {
        long nextWork = SaturatingAdd(_intersectionWork, Math.Max(0L, addedWork));
        if (nextWork > PdfPageClipPath.MaximumTextClippingIntersectionWork) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.TextClippingIntersectionWork,
                PdfPageClipPath.MaximumTextClippingIntersectionWork,
                nextWork);
        }
        _intersectionWork = nextWork;
    }

    private static long CountVertices(IReadOnlyList<List<OfficePoint>> contours) {
        long vertices = 0L;
        for (int index = 0; index < contours.Count; index++) {
            vertices = SaturatingAdd(vertices, contours[index].Count);
        }
        return vertices;
    }

    private static long SaturatingAdd(long left, long right) =>
        right > long.MaxValue - left ? long.MaxValue : left + right;

}
