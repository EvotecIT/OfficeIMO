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
        if (activeClipPath.HasValue && !activeClipPath.Value.IsRectangle && !clipPath.IsRectangle) {
            long activeContours = CountContours(activeClipPath.Value.Commands);
            long nextContours = CountContours(clipPath.Commands);
            long overlapChecks = nextContours * Math.Max(0L, nextContours - 1L) / 2L;
            long intersectionChecks = activeContours * nextContours;
            long nextWork = checked(_intersectionWork + overlapChecks + intersectionChecks);
            if (nextWork > PdfPageClipPath.MaximumTextClippingIntersectionWork) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.TextClippingIntersectionWork,
                    PdfPageClipPath.MaximumTextClippingIntersectionWork,
                    nextWork);
            }
            _intersectionWork = nextWork;
        }

        return PdfPageClipPath.ResolveActiveClip(activeClipPath, clipPath);
    }

    private static long CountContours(IReadOnlyList<OfficePathCommand> commands) {
        long contours = 0L;
        for (int index = 0; index < commands.Count; index++) {
            if (commands[index].Kind == OfficePathCommandKind.MoveTo) contours++;
        }
        return contours;
    }
}
