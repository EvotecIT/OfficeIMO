namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private readonly record struct RichTextInterval(double X, double Width);
    private sealed record RichTextRowFrame(IReadOnlyList<RichTextInterval> Intervals, double Gap);

    /// <summary>A segment range and its available interval on one physical text row.</summary>
    private sealed record RichLineFragment(int Start, int Count, double Origin, double Width);

    /// <summary>
    /// Tracks available intervals for the existing breaker. Moving to another interval does not
    /// consume source text or advance a baseline; only completing a physical row advances height.
    /// </summary>
    private sealed class RichRowComposition {
        private readonly Func<int, double, double, double, RichTextRowFrame> resolve;
        private readonly List<(int Row, double Origin, double Width, double Gap)> placements = new();
        private RichTextRowFrame? frame;
        private int interval;
        private int row;

        internal RichRowComposition(Func<int, double, double, double, RichTextRowFrame> resolve) => this.resolve = resolve;
        internal int RowIndex => row;

        internal (double Width, double Origin, double Gap) Resolve(int fragment, double completedHeight,
            double height, double minimumWidth, bool empty, out bool invalidated) {
            invalidated = false;
            RichTextRowFrame candidate = resolve(row, completedHeight, height, empty ? minimumWidth : 0D);
            if (candidate.Intervals.Count == 0) throw new InvalidOperationException("A text row must have an available interval after clearance.");
            if (frame != null && interval > 0 && !SamePreviousIntervals(frame, candidate, interval)) {
                invalidated = true;
                return Current(fragment);
            }
            frame = candidate;
            interval = Math.Min(interval, frame.Intervals.Count - 1);
            if (empty && minimumWidth > 0D)
                while (interval + 1 < frame.Intervals.Count && frame.Intervals[interval].Width + RichTextWidthTolerance < minimumWidth)
                    interval++;
            if (empty && minimumWidth > 0D && interval > 0 &&
                frame.Intervals[interval].Width + RichTextWidthTolerance < minimumWidth) {
                invalidated = true;
                return Current(fragment);
            }
            return Current(fragment);
        }

        private static bool SamePreviousIntervals(RichTextRowFrame previous, RichTextRowFrame candidate, int count) {
            if (previous.Gap != candidate.Gap || candidate.Intervals.Count < count) return false;
            for (int index = 0; index < count; index++)
                if (previous.Intervals[index] != candidate.Intervals[index]) return false;
            return true;
        }

        private (double Width, double Origin, double Gap) Current(int fragment) {
            RichTextInterval slot = frame!.Intervals[interval];
            while (placements.Count <= fragment) placements.Add(default);
            placements[fragment] = (row, slot.X, slot.Width, frame.Gap);
            return (slot.Width, slot.X, frame.Gap);
        }

        internal bool TryContinue(int nextFragment, bool forceRow) {
            if (!forceRow && frame != null && interval + 1 < frame.Intervals.Count) {
                interval++;
                Current(nextFragment);
                return true;
            }
            row++;
            interval = 0;
            frame = null;
            return false;
        }

        /// <summary>Collapses interval fragments into rows before pagination observes any line count.</summary>
        internal (List<List<RichSeg>> Lines, List<double> Heights) Complete(
            List<List<RichSeg>> fragments, List<double> advances, PdfStandardFont baseFont,
            double fontSize, PdfCreateOptions? options, PdfLineSpacing? spacing) {
            var rows = new List<List<RichSeg>>();
            var heights = new List<double>();
            for (int start = 0; start < fragments.Count;) {
                int end = start + 1;
                while (end < fragments.Count && placements[end].Row == placements[start].Row) end++;
                var first = (RichLine)fragments[start];
                RichLine line = first.CopyPlacement();
                line.IsFragmentContext = true;
                line.FragmentTopGap = placements[start].Gap;
                line.BarOriginOffset = 0D;
                var ranges = new List<RichLineFragment>();
                double advance = 0D;
                for (int index = start; index < end; index++) {
                    var placement = placements[index];
                    if (fragments[index].Count > 0 || end == start + 1) {
                        ranges.Add(new RichLineFragment(line.Count, fragments[index].Count, placement.Origin, placement.Width));
                        line.AddRange(fragments[index]);
                    }
                    advance += advances[index];
                }
                line.Fragments = ranges.AsReadOnly();
                line.BarLineHeight = Math.Max(0D, advance - line.FragmentTopGap);
                RefreshRichLinePlacementMetrics(line, baseFont, fontSize, options, spacing);
                rows.Add(line);
                heights.Add(advance);
                start = end;
            }
            return (rows, heights);
        }
    }

    private sealed record RichPaintRows(List<List<RichSeg>> Lines, List<double> Heights,
        List<double> Offsets, List<double> Widths, List<double> Gaps);

    /// <summary>
    /// Adapts retained row plans to the ordinary painter without another layout or glyph pass.
    /// Zero advances exist only in this call-local paint view; pagination sees physical rows.
    /// </summary>
    private static RichPaintRows? ExpandRichPaintRows(List<List<RichSeg>> lines,
        IReadOnlyList<double> heights, double fallbackLeading, double origin, double width,
        double? firstX, double? firstWidth, IReadOnlyList<double>? offsets,
        IReadOnlyList<double>? widths, IReadOnlyList<double>? gaps) {
        if (!lines.Any(line => line is RichLine { Fragments: not null })) return null;
        var painted = new List<List<RichSeg>>();
        var advances = new List<double>();
        var paintedOffsets = new List<double>();
        var paintedWidths = new List<double>();
        var paintedGaps = new List<double>();
        for (int index = 0; index < lines.Count; index++) {
            double advance = GetRichLineHeight(heights, index, fallbackLeading);
            if (lines[index] is RichLine { Fragments: { } fragments } row) {
                for (int part = 0; part < fragments.Count; part++) {
                    RichLineFragment fragment = fragments[part];
                    RichLine paintedLine = row.CopyPlacement();
                    paintedLine.IsFragmentContext = true;
                    paintedLine.AddRange(row.GetRange(fragment.Start, fragment.Count));
                    // Stationary paragraph rules belong to the row, not every fragment.
                    if (part != 0) paintedLine.BarTabStops = null;
                    else paintedLine.BarOriginOffset = -fragment.Origin;
                    painted.Add(paintedLine);
                    advances.Add(part == fragments.Count - 1 ? advance : 0D);
                    paintedOffsets.Add(fragment.Origin);
                    paintedWidths.Add(fragment.Width);
                    paintedGaps.Add(row.FragmentTopGap);
                }
            } else {
                painted.Add(lines[index]);
                advances.Add(advance);
                paintedOffsets.Add(ResolveRichLineXOrigin(origin, firstX, offsets, index) - origin);
                paintedWidths.Add(ResolveRichLineWidth(width, firstWidth, widths, index));
                paintedGaps.Add(gaps?[index] ?? 0D);
            }
        }
        return new RichPaintRows(painted, advances, paintedOffsets, paintedWidths, paintedGaps);
    }
}
