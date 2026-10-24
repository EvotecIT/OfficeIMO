namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly record struct TextExclusionBounds(double Left, double Right, double Top,
        double Bottom, PdfImageTextWrapKind Kind);

    /// <summary>Subtracts one exclusion without selecting or fitting a text interval.</summary>
    private static List<RichTextInterval> SubtractTextExclusion(IReadOnlyList<RichTextInterval> intervals,
        double left, double right) {
        var remaining = new List<RichTextInterval>();
        foreach (RichTextInterval interval in intervals) {
            double end = interval.X + interval.Width;
            if (right <= interval.X || left >= end) { remaining.Add(interval); continue; }
            if (left > interval.X) remaining.Add(new RichTextInterval(interval.X, Math.Min(end, left) - interval.X));
            if (right < end) remaining.Add(new RichTextInterval(Math.Max(interval.X, right), end - Math.Max(interval.X, right)));
        }
        return remaining;
    }

    private sealed partial class LayoutContext {
        private IEnumerable<TextExclusionBounds> GetCurrentTextExclusions() {
            foreach (var table in floatingTables)
                if (ReferenceEquals(table.Page, currentPage))
                    yield return new TextExclusionBounds(table.Left, table.Right, table.Top, table.Bottom, PdfImageTextWrapKind.Square);
            foreach (PictureAnchorPlacement picture in CurrentPictureExclusions()) {
                PdfImageTextExclusion wrap = picture.Exclusion!;
                if (wrap.Kind == PdfImageTextWrapKind.Square && wrap.Side != PdfImageTextWrapSide.BothSides)
                    throw new NotSupportedException("Only both-side rectangular picture wrapping is currently supported.");
                yield return new TextExclusionBounds(picture.X - wrap.DistanceLeft,
                    picture.X + picture.Width + wrap.DistanceRight, picture.Top + wrap.DistanceTop,
                    picture.Top - picture.Height - wrap.DistanceBottom, wrap.Kind);
            }
        }

        /// <summary>
        /// Returns all positive intervals for the occupied row box. A source word that fits the
        /// unobstructed frame may clear to an actual exclusion boundary; no font-based width floor applies.
        /// </summary>
        private static RichTextRowFrame ResolveTextRowFrame(double left, double width, double top,
            double height, double minimumWidth, IReadOnlyList<TextExclusionBounds> exclusions) {
            double originalTop = top;
            for (int attempt = 0; attempt <= exclusions.Count; attempt++) {
                var intervals = new List<RichTextInterval> { new(left, width) };
                double nextBoundary = double.NegativeInfinity;
                foreach (TextExclusionBounds region in exclusions) {
                    if (top - height >= region.Top || top <= region.Bottom) continue;
                    if (region.Kind != PdfImageTextWrapKind.TopAndBottom &&
                        (region.Right <= left || region.Left >= left + width)) continue;
                    nextBoundary = Math.Max(nextBoundary, region.Bottom);
                    intervals = region.Kind == PdfImageTextWrapKind.TopAndBottom
                        ? new List<RichTextInterval>()
                        : SubtractTextExclusion(intervals, region.Left, region.Right);
                }
                if (intervals.Count > 0 && (minimumWidth <= 0D || minimumWidth > width ||
                    intervals.Any(slot => slot.Width + RichTextWidthTolerance >= minimumWidth)))
                    return new RichTextRowFrame(intervals, originalTop - top);
                if (double.IsNegativeInfinity(nextBoundary) || nextBoundary >= top)
                    throw new InvalidOperationException("Picture exclusion clearance could not advance.");
                top = nextBoundary;
            }
            throw new InvalidOperationException("Picture exclusion clearance did not resolve at a finite source boundary.");
        }
    }
}