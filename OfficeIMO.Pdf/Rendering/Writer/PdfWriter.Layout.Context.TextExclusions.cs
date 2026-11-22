namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// The area text keeps clear of, in page points. A polygon wrap also carries its exclusion and the
    /// picture frame, offset from Left and Top, so moving Top and Bottom moves the polygon with them.
    /// </summary>
    private readonly record struct TextExclusionBounds(double Left, double Right, double Top,
        double Bottom, PdfImageTextWrapKind Kind, PdfImageTextExclusion? Wrap = null,
        double PictureLeftOffset = 0D, double PictureTopOffset = 0D, double PictureWidth = 0D, double PictureHeight = 0D);

    /// <summary>
    /// The horizontal extent of a wrap polygon over a band of rows, widened by the wrap distances as
    /// Word does: the band grows by the top and bottom distances, the extent by the side distances.
    /// </summary>
    private static (double Left, double Right)? GetPolygonTextExclusion(TextExclusionBounds region,
        double bandBottom, double bandTop) {
        PdfImageTextExclusion wrap = region.Wrap!;
        IReadOnlyList<(double X, double Y)> polygon = wrap.Polygon!;
        double low = bandBottom - wrap.DistanceTop, high = bandTop + wrap.DistanceBottom;
        double originX = region.Left + region.PictureLeftOffset, originY = region.Top - region.PictureTopOffset;
        double left = double.PositiveInfinity, right = double.NegativeInfinity;
        for (int index = 0; index < polygon.Count; index++) {
            (double ax, double ay) = Point(polygon[index]);
            (double bx, double by) = Point(polygon[(index + 1) % polygon.Count]);
            if (Math.Max(ay, by) < low || Math.Min(ay, by) > high) continue;
            if (ay == by) {
                Include(ax); Include(bx);
                continue;
            }
            // Clip the edge to the band and keep the x range of what remains.
            double y0 = Math.Max(low, Math.Min(ay, by)), y1 = Math.Min(high, Math.Max(ay, by));
            Include(ax + (bx - ax) * (y0 - ay) / (by - ay));
            Include(ax + (bx - ax) * (y1 - ay) / (by - ay));
        }
        if (left > right) return null;
        return (left - wrap.DistanceLeft, right + wrap.DistanceRight);

        (double X, double Y) Point((double X, double Y) point) =>
            (originX + point.X * region.PictureWidth, originY - point.Y * region.PictureHeight);
        void Include(double x) { left = Math.Min(left, x); right = Math.Max(right, x); }
    }

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
                if (wrap.Kind != PdfImageTextWrapKind.TopAndBottom && wrap.Side != PdfImageTextWrapSide.BothSides)
                    throw new NotSupportedException("Only both-side picture wrapping is currently supported.");
                if (wrap.Polygon is { } polygon) {
                    // The bounds hold the whole polygon widened by the wrap distances; rows test the polygon itself.
                    double minX = picture.X + polygon.Min(point => point.X) * picture.Width;
                    double maxX = picture.X + polygon.Max(point => point.X) * picture.Width;
                    double maxY = picture.Top - polygon.Min(point => point.Y) * picture.Height;
                    double minY = picture.Top - polygon.Max(point => point.Y) * picture.Height;
                    double left = minX - wrap.DistanceLeft, top = maxY + wrap.DistanceTop;
                    yield return new TextExclusionBounds(left, maxX + wrap.DistanceRight, top,
                        minY - wrap.DistanceBottom, wrap.Kind, wrap, picture.X - left, top - picture.Top,
                        picture.Width, picture.Height);
                    continue;
                }
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
            // A polygon can narrow at every vertex, so it may need a step for each one.
            int attempts = exclusions.Count + 1;
            foreach (TextExclusionBounds region in exclusions) attempts += 2 * (region.Wrap?.Polygon?.Count ?? 0);
            double? blockedByPolygon = null;
            for (int attempt = 0; attempt < attempts; attempt++) {
                if (TryFitTextRow(left, width, top, height, minimumWidth, exclusions,
                    out double nextBoundary, out bool polygonInTheWay) is { } intervals) {
                    // The row resumes as soon as it fits rather than at the step.
                    if (blockedByPolygon is { } blocked) top = ResumeTextRow(blocked, top, ref intervals);
                    return new RichTextRowFrame(intervals, originalTop - top);
                }
                if (double.IsNegativeInfinity(nextBoundary) || nextBoundary >= top)
                    throw new InvalidOperationException("Picture exclusion clearance could not advance.");
                blockedByPolygon = null;
                if (polygonInTheWay) {
                    // Between two steps the extent changes linearly, but the edge that sets it can change
                    // partway, so a concave polygon can leave room between two blocked steps. Probe the
                    // interval in small steps before passing it.
                    int probes = Math.Min(MaximumPolygonRowProbes, Math.Max(1, (int)Math.Ceiling((top - nextBoundary) / PolygonRowProbeStep)));
                    double step = (top - nextBoundary) / probes, last = top;
                    for (int probe = 1; probe < probes; probe++) {
                        double candidate = top - probe * step;
                        if (TryFitTextRow(left, width, candidate, height, minimumWidth, exclusions, out _, out _) is { } found) {
                            double resumed = ResumeTextRow(last, candidate, ref found);
                            return new RichTextRowFrame(found, originalTop - resumed);
                        }
                        last = candidate;
                    }
                    blockedByPolygon = last;
                }
                top = nextBoundary;
            }
            throw new InvalidOperationException("Picture exclusion clearance did not resolve at a finite source boundary.");

            // Bisects between a blocked top and a lower one that fits, to the first top where the row fits.
            double ResumeTextRow(double blocked, double fits, ref List<RichTextInterval> found) {
                for (int step = 0; step < 48 && blocked - fits > .001D; step++) {
                    double middle = (fits + blocked) / 2D;
                    if (TryFitTextRow(left, width, middle, height, minimumWidth, exclusions, out _, out _) is { } fitted) {
                        fits = middle;
                        found = fitted;
                    } else blocked = middle;
                }
                return fits;
            }
        }

        /// <summary>
        /// The spacing of the probes between two polygon steps, with a cap per interval so a tall picture
        /// stays bounded; room shorter than a probe step can be missed.
        /// </summary>
        private const double PolygonRowProbeStep = .25D;
        private const int MaximumPolygonRowProbes = 256;

        /// <summary>
        /// Returns the row's free intervals when one fits, else null with the next row top worth trying:
        /// the bottom of a rectangular exclusion, or the next point where a wrap polygon's extent changes.
        /// </summary>
        private static List<RichTextInterval>? TryFitTextRow(double left, double width, double top, double height,
            double minimumWidth, IReadOnlyList<TextExclusionBounds> exclusions, out double nextBoundary, out bool polygonInTheWay) {
            var intervals = new List<RichTextInterval> { new(left, width) };
            nextBoundary = double.NegativeInfinity;
            polygonInTheWay = false;
            foreach (TextExclusionBounds region in exclusions) {
                if (top - height >= region.Top || top <= region.Bottom) continue;
                if (region.Kind != PdfImageTextWrapKind.TopAndBottom &&
                    (region.Right <= left || region.Left >= left + width)) continue;
                if (region.Wrap?.Polygon != null) {
                    if (GetPolygonTextExclusion(region, top - height, top) is not { } span ||
                        span.Right <= left || span.Left >= left + width) continue;
                    polygonInTheWay = true;
                    nextBoundary = Math.Max(nextBoundary, GetNextPolygonRowTop(region, top, height));
                    intervals = SubtractTextExclusion(intervals, span.Left, span.Right);
                    continue;
                }
                nextBoundary = Math.Max(nextBoundary, region.Bottom);
                intervals = region.Kind == PdfImageTextWrapKind.TopAndBottom
                    ? new List<RichTextInterval>()
                    : SubtractTextExclusion(intervals, region.Left, region.Right);
            }
            return intervals.Count > 0 && (minimumWidth <= 0D || minimumWidth > width ||
                intervals.Any(slot => slot.Width + RichTextWidthTolerance >= minimumWidth)) ? intervals : null;
        }

        /// <summary>
        /// The highest row top below the current one at which a polygon vertex enters or leaves the row
        /// widened by the wrap distances, or the polygon's bottom.
        /// </summary>
        private static double GetNextPolygonRowTop(TextExclusionBounds region, double top, double height) {
            PdfImageTextExclusion wrap = region.Wrap!;
            double originY = region.Top - region.PictureTopOffset;
            double next = region.Bottom;
            foreach ((double _, double y) in wrap.Polygon!) {
                double vertex = originY - y * region.PictureHeight;
                foreach (double candidate in new[] { vertex - wrap.DistanceBottom, vertex + height + wrap.DistanceTop })
                    if (candidate < top - .001D) next = Math.Max(next, candidate);
            }
            return next;
        }
    }
}