namespace OfficeIMO.Pdf;

/// <summary>
/// Orders page regions by recursively partitioning whitespace into horizontal bands and vertical columns.
/// </summary>
internal sealed class PdfRecursiveXyCutReadingOrderStage : IPdfReadingOrderStage {
    private const int MaximumDepth = 64;

    /// <inheritdoc />
    public IReadOnlyList<PdfUnderstandingRegion> Order(
        PdfUnderstandingPageContext context,
        IReadOnlyList<PdfUnderstandingRegion> regions) {
        Guard.NotNull(context, nameof(context));
        Guard.NotNull(regions, nameof(regions));
        if (regions.Count <= 1) {
            return regions.ToArray();
        }

        RegionBox[] boxes = regions.Select(static region => RegionBox.From(region)).ToArray();
        double medianFontSize = Median(boxes.Select(static box => box.FontSize));
        double minimumHorizontalGap = Math.Max(6D, medianFontSize * 0.8D);
        double minimumVerticalGap = Math.Max(context.LayoutOptions.MinGutterWidth, medianFontSize * 1.5D);
        var ordered = new List<PdfUnderstandingRegion>(regions.Count);

        AppendPartition(
            boxes,
            ordered,
            minimumHorizontalGap,
            minimumVerticalGap,
            context.LayoutOptions.ForceSingleColumn,
            depth: 0);

        return ordered.ToArray();
    }

    private static void AppendPartition(
        IReadOnlyList<RegionBox> boxes,
        List<PdfUnderstandingRegion> ordered,
        double minimumHorizontalGap,
        double minimumVerticalGap,
        bool forceSingleColumn,
        int depth) {
        if (boxes.Count == 0) {
            return;
        }

        if (boxes.Count == 1) {
            ordered.Add(boxes[0].Region);
            return;
        }

        if (depth >= MaximumDepth) {
            AppendFallback(boxes, ordered);
            return;
        }

        WhitespaceCut? horizontal = FindBestCut(boxes, horizontal: true, minimumHorizontalGap);
        WhitespaceCut? vertical = forceSingleColumn
            ? null
            : FindBestCut(boxes, horizontal: false, minimumVerticalGap);
        if (!forceSingleColumn &&
            !vertical.HasValue &&
            TryAppendSpanningEdgeRegion(
                boxes,
                ordered,
                minimumHorizontalGap,
                minimumVerticalGap,
                depth)) {
            return;
        }
        WhitespaceCut? selected = SelectCut(boxes, horizontal, vertical);
        if (!selected.HasValue) {
            AppendFallback(boxes, ordered);
            return;
        }

        WhitespaceCut cut = selected.Value;
        var first = new List<RegionBox>();
        var second = new List<RegionBox>();
        for (int index = 0; index < boxes.Count; index++) {
            RegionBox box = boxes[index];
            double center = cut.Horizontal
                ? (box.Bottom + box.Top) / 2D
                : (box.Left + box.Right) / 2D;
            if (center < cut.Midpoint) {
                first.Add(box);
            } else {
                second.Add(box);
            }
        }

        if (first.Count == 0 || second.Count == 0) {
            AppendFallback(boxes, ordered);
            return;
        }

        if (cut.Horizontal) {
            AppendPartition(second, ordered, minimumHorizontalGap, minimumVerticalGap, forceSingleColumn, depth + 1);
            AppendPartition(first, ordered, minimumHorizontalGap, minimumVerticalGap, forceSingleColumn, depth + 1);
        } else {
            AppendPartition(first, ordered, minimumHorizontalGap, minimumVerticalGap, forceSingleColumn, depth + 1);
            AppendPartition(second, ordered, minimumHorizontalGap, minimumVerticalGap, forceSingleColumn, depth + 1);
        }
    }

    private static WhitespaceCut? SelectCut(
        IReadOnlyList<RegionBox> boxes,
        WhitespaceCut? horizontal,
        WhitespaceCut? vertical) {
        if (!vertical.HasValue) return horizontal;
        if (!horizontal.HasValue) return vertical;
        return HasVerticalOverlapAcrossCut(boxes, vertical.Value) ? vertical : horizontal;
    }

    private static bool HasVerticalOverlapAcrossCut(IReadOnlyList<RegionBox> boxes, WhitespaceCut verticalCut) {
        var firstSide = new List<Interval>();
        var secondSide = new List<Interval>();
        for (int index = 0; index < boxes.Count; index++) {
            RegionBox box = boxes[index];
            var interval = new Interval(box.Bottom, box.Top);
            if ((box.Left + box.Right) / 2D < verticalCut.Midpoint) {
                firstSide.Add(interval);
            } else {
                secondSide.Add(interval);
            }
        }

        firstSide.Sort(static (first, second) => first.Start.CompareTo(second.Start));
        secondSide.Sort(static (first, second) => first.Start.CompareTo(second.Start));
        int firstIndex = 0;
        int secondIndex = 0;
        while (firstIndex < firstSide.Count && secondIndex < secondSide.Count) {
            Interval first = firstSide[firstIndex];
            Interval second = secondSide[secondIndex];
            if (Math.Min(first.End, second.End) > Math.Max(first.Start, second.Start)) {
                return true;
            }

            if (first.End <= second.End) {
                firstIndex++;
            } else {
                secondIndex++;
            }
        }

        return false;
    }

    private static WhitespaceCut? FindBestCut(
        IReadOnlyList<RegionBox> boxes,
        bool horizontal,
        double minimumGap) {
        Interval[] intervals = boxes
            .Select(box => horizontal
                ? new Interval(box.Bottom, box.Top)
                : new Interval(box.Left, box.Right))
            .OrderBy(static interval => interval.Start)
            .ThenBy(static interval => interval.End)
            .ToArray();
        if (intervals.Length < 2) {
            return null;
        }

        double occupiedEnd = intervals[0].End;
        WhitespaceCut? best = null;
        for (int index = 1; index < intervals.Length; index++) {
            Interval interval = intervals[index];
            double gap = interval.Start - occupiedEnd;
            if (gap >= minimumGap && (!best.HasValue || gap > best.Value.Size)) {
                best = new WhitespaceCut(horizontal, occupiedEnd, interval.Start);
            }
            occupiedEnd = Math.Max(occupiedEnd, interval.End);
        }

        return best;
    }

    private static bool TryAppendSpanningEdgeRegion(
        IReadOnlyList<RegionBox> boxes,
        List<PdfUnderstandingRegion> ordered,
        double minimumHorizontalGap,
        double minimumVerticalGap,
        int depth) {
        foreach (RegionBox candidate in boxes.OrderByDescending(static box => box.Top)) {
            RegionBox[] remaining = boxes.Where(box => !ReferenceEquals(box.Region, candidate.Region)).ToArray();
            if (remaining.Length < 2) continue;

            double candidateCenter = (candidate.Bottom + candidate.Top) / 2D;
            bool beforeRemaining = remaining.All(box => (box.Bottom + box.Top) / 2D < candidateCenter);
            bool afterRemaining = remaining.All(box => (box.Bottom + box.Top) / 2D > candidateCenter);
            if (!beforeRemaining && !afterRemaining) continue;

            WhitespaceCut? columnCut = FindBestCut(remaining, horizontal: false, minimumVerticalGap);
            if (!columnCut.HasValue ||
                candidate.Left > columnCut.Value.Start ||
                candidate.Right < columnCut.Value.End) {
                continue;
            }

            if (beforeRemaining) ordered.Add(candidate.Region);
            AppendPartition(remaining, ordered, minimumHorizontalGap, minimumVerticalGap, forceSingleColumn: false, depth + 1);
            if (afterRemaining) ordered.Add(candidate.Region);
            return true;
        }

        return false;
    }

    private static void AppendFallback(IReadOnlyList<RegionBox> boxes, List<PdfUnderstandingRegion> ordered) {
        foreach (RegionBox box in boxes
                     .OrderByDescending(static box => box.Top)
                     .ThenBy(static box => box.Left)
                     .ThenByDescending(static box => box.Bottom)) {
            ordered.Add(box.Region);
        }
    }

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(static value => value).ToArray();
        if (ordered.Length == 0) {
            return 0D;
        }
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0
            ? (ordered[middle - 1] + ordered[middle]) / 2D
            : ordered[middle];
    }

    private readonly struct RegionBox {
        private RegionBox(
            PdfUnderstandingRegion region,
            double left,
            double right,
            double bottom,
            double top,
            double fontSize) {
            Region = region;
            Left = left;
            Right = right;
            Bottom = bottom;
            Top = top;
            FontSize = fontSize;
        }

        internal PdfUnderstandingRegion Region { get; }
        internal double Left { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal double Top { get; }
        internal double FontSize { get; }

        internal static RegionBox From(PdfUnderstandingRegion region) {
            (double left, double right, double bottom, double top, double fontSize) = GetSourceBounds(region);
            return new RegionBox(
                region,
                left,
                right,
                bottom,
                top,
                fontSize);
        }
    }

    internal static (double Left, double Right, double Bottom, double Top, double FontSize) GetSourceBounds(
        PdfUnderstandingRegion region) {
        double fontSize = Math.Max(1D, region.Lines.Max(static line => line.FontSize));
        double left = region.XStart;
        double right = Math.Max(region.XStart, region.XEnd);
        double bottom = region.YBottom - (fontSize * 0.25D);
        double top = region.YTop + (fontSize * 0.8D);

        foreach (PdfUnderstandingWord word in region.Lines.SelectMany(static line => line.Words)) {
            PdfTextSpan? sourceRun = word.SourceRuns.Count > 0 ? word.SourceRuns[0] : null;
            double wordFontSize = Math.Max(1D, word.FontSize);
            double radians = word.RotationDegrees * Math.PI / 180D;
            double alongX = Math.Cos(radians);
            double alongY = Math.Sin(radians);
            double normalX = -alongY;
            double normalY = alongX;
            double advance = GetWordAdvance(word, sourceRun, alongX, wordFontSize);
            double startX = alongX >= 0D ? word.XStart : word.XEnd;
            double startY = word.BaselineY;
            double endX = startX + (alongX * advance);
            double endY = startY + (alongY * advance);
            ExpandBounds(startX, startY, normalX, normalY, wordFontSize, ref left, ref right, ref bottom, ref top);
            ExpandBounds(endX, endY, normalX, normalY, wordFontSize, ref left, ref right, ref bottom, ref top);
        }

        return (left, right, bottom, top, fontSize);
    }

    private static double GetWordAdvance(
        PdfUnderstandingWord word,
        PdfTextSpan? sourceRun,
        double alongX,
        double fontSize) {
        double horizontalExtent = Math.Abs(word.XEnd - word.XStart);
        if (horizontalExtent > 0.001D && Math.Abs(alongX) > 0.05D) {
            return horizontalExtent / Math.Abs(alongX);
        }

        if (sourceRun is not null && sourceRun.Advance > 0D && !string.IsNullOrEmpty(sourceRun.Text)) {
            return sourceRun.Advance * word.Text.Length / sourceRun.Text.Length;
        }

        return word.Text.Length * fontSize * 0.55D;
    }

    private static void ExpandBounds(
        double x,
        double y,
        double normalX,
        double normalY,
        double fontSize,
        ref double left,
        ref double right,
        ref double bottom,
        ref double top) {
        ExpandPoint(x - (normalX * fontSize * 0.25D), y - (normalY * fontSize * 0.25D), ref left, ref right, ref bottom, ref top);
        ExpandPoint(x + (normalX * fontSize * 0.8D), y + (normalY * fontSize * 0.8D), ref left, ref right, ref bottom, ref top);
    }

    private static void ExpandPoint(
        double x,
        double y,
        ref double left,
        ref double right,
        ref double bottom,
        ref double top) {
        left = Math.Min(left, x);
        right = Math.Max(right, x);
        bottom = Math.Min(bottom, y);
        top = Math.Max(top, y);
    }

    private readonly struct Interval {
        internal Interval(double start, double end) {
            Start = start;
            End = end;
        }

        internal double Start { get; }
        internal double End { get; }
    }

    private readonly struct WhitespaceCut {
        internal WhitespaceCut(bool horizontal, double start, double end) {
            Horizontal = horizontal;
            Start = start;
            End = end;
        }

        internal bool Horizontal { get; }
        internal double Start { get; }
        internal double End { get; }
        internal double Size => End - Start;
        internal double Midpoint => (Start + End) / 2D;
    }
}
