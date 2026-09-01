namespace OfficeIMO.Pdf;

/// <summary>Semantic content categories ordered for editable and review-surface reconstruction.</summary>
public enum PdfLogicalReadingOrderKind {
    /// <summary>Line-level text not otherwise consumed by a semantic projection.</summary>
    TextBlock,
    /// <summary>Heading inferred from a logical text block.</summary>
    Heading,
    /// <summary>Paragraph inferred from one or more logical text blocks.</summary>
    Paragraph,
    /// <summary>List item inferred from a logical text block.</summary>
    ListItem,
    /// <summary>Detected logical table.</summary>
    Table,
    /// <summary>Placed image resource.</summary>
    Image,
    /// <summary>Link annotation.</summary>
    Link,
    /// <summary>AcroForm widget annotation.</summary>
    FormWidget
}

/// <summary>Controls which logical page content participates in shared reading-order analysis.</summary>
public enum PdfLogicalReadingOrderScope {
    /// <summary>Orders the semantic body while excluding classified running headers and footers.</summary>
    SemanticBody,
    /// <summary>Orders all page content, including classified running headers and footers.</summary>
    PageContent
}

/// <summary>
/// One logical page item in crop-, rotation-, and column-aware reading order.
/// </summary>
public sealed class PdfLogicalReadingOrderItem {
    internal PdfLogicalReadingOrderItem(
        PdfLogicalReadingOrderKind kind,
        int sourceIndex,
        int placementIndex,
        int orderIndex,
        int columnIndex,
        bool spansColumns,
        bool hasGeometry,
        bool isClipped,
        double left,
        double top,
        double right,
        double bottom,
        double confidence,
        IReadOnlyList<PdfInferenceEvidence> evidence) {
        Kind = kind;
        SourceIndex = sourceIndex;
        PlacementIndex = placementIndex;
        OrderIndex = orderIndex;
        ColumnIndex = columnIndex;
        SpansColumns = spansColumns;
        HasGeometry = hasGeometry;
        IsClipped = isClipped;
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
        Confidence = PdfInference.Clamp(confidence);
        Evidence = evidence;
    }

    /// <summary>Semantic source collection containing the item.</summary>
    public PdfLogicalReadingOrderKind Kind { get; }

    /// <summary>Zero-based index in the matching logical page collection.</summary>
    public int SourceIndex { get; }

    /// <summary>Zero-based image placement index, or -1 when the item is not a placed image.</summary>
    public int PlacementIndex { get; }

    /// <summary>Zero-based position in inferred reading order.</summary>
    public int OrderIndex { get; }

    /// <summary>Zero-based inferred column within the current page band.</summary>
    public int ColumnIndex { get; }

    /// <summary>True when the item is wide enough to divide column bands.</summary>
    public bool SpansColumns { get; }

    /// <summary>True when usable source geometry was available.</summary>
    public bool HasGeometry { get; }

    /// <summary>True when source geometry extended beyond the visible crop boundary.</summary>
    public bool IsClipped { get; }

    /// <summary>Left edge in top-left visual page coordinates.</summary>
    public double Left { get; }

    /// <summary>Top edge in top-left visual page coordinates.</summary>
    public double Top { get; }

    /// <summary>Right edge in top-left visual page coordinates.</summary>
    public double Right { get; }

    /// <summary>Bottom edge in top-left visual page coordinates.</summary>
    public double Bottom { get; }

    /// <summary>Normalized inference confidence from 0 to 1.</summary>
    public double Confidence { get; }

    /// <summary>Stable evidence supporting the inferred position.</summary>
    public IReadOnlyList<PdfInferenceEvidence> Evidence { get; }
}

/// <summary>Shared logical reading-order analysis for reverse-conversion adapters.</summary>
public static class PdfLogicalReadingOrderAnalysis {
    private const double SpanningWidthRatio = 0.62D;

    /// <summary>
    /// Orders semantic page items using visible crop geometry, page rotation, spanning bands, and columns.
    /// </summary>
    public static IReadOnlyList<PdfLogicalReadingOrderItem> Analyze(PdfLogicalPage page) =>
        Analyze(page, PdfLogicalReadingOrderScope.SemanticBody);

    /// <summary>
    /// Orders logical page items using visible crop geometry, page rotation, spanning bands, and columns.
    /// </summary>
    /// <param name="page">Logical page to analyze.</param>
    /// <param name="scope">Whether to return semantic body content or all page content.</param>
    public static IReadOnlyList<PdfLogicalReadingOrderItem> Analyze(
        PdfLogicalPage page,
        PdfLogicalReadingOrderScope scope) {
        Guard.NotNull(page, nameof(page));
        if (scope is not (PdfLogicalReadingOrderScope.SemanticBody or PdfLogicalReadingOrderScope.PageContent)) {
            throw new ArgumentOutOfRangeException(nameof(scope), scope, "Unsupported logical reading-order scope.");
        }
        Action<long>? consumeWork = page.Analysis.ConsumeWork;
        Action? cancellationCheck = page.Analysis.CancellationCheck;
        cancellationCheck?.Invoke();
        var candidates = BuildCandidates(page, scope);
        if (candidates.Count == 0) return Array.Empty<PdfLogicalReadingOrderItem>();

        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        var positioned = candidates.Where(static item => item.HasGeometry).ToArray();
        var unpositioned = candidates.Where(static item => !item.HasGeometry).OrderBy(static item => item.Sequence).ToArray();
        double[] pageColumnAnchors = FindRepeatedColumnAnchors(
            positioned.Where(item => item.Right - item.Left < Math.Max(1D, pageWidth) * SpanningWidthRatio),
            pageWidth,
            consumeWork,
            cancellationCheck);
        Candidate[] spanning = positioned
            .Where(item => item.Right - item.Left >= Math.Max(1D, pageWidth) * SpanningWidthRatio ||
                IsCenteredBandDivider(item, pageColumnAnchors, pageWidth))
            .OrderBy(static item => item.Top)
            .ThenBy(static item => item.Left)
            .ToArray();
        var ordered = new List<Candidate>(candidates.Count);
        var consumed = new HashSet<Candidate>();
        double bandTop = 0D;
        for (int index = 0; index < spanning.Length; index++) {
            Candidate divider = spanning[index];
            AddBand(positioned.Where(item => !consumed.Contains(item) && !ReferenceEquals(item, divider) && item.Top >= bandTop && item.Top < divider.Top), ordered, consumed);
            divider.SpansColumns = true;
            divider.ColumnIndex = 0;
            if (consumed.Add(divider)) ordered.Add(divider);
            bandTop = Math.Max(bandTop, divider.Bottom);
        }
        AddBand(positioned.Where(item => !consumed.Contains(item) && item.Top >= bandTop), ordered, consumed);
        ordered.AddRange(unpositioned);

        var result = new PdfLogicalReadingOrderItem[ordered.Count];
        for (int index = 0; index < ordered.Count; index++) {
            Candidate item = ordered[index];
            double confidence = item.HasGeometry ? 0.9D : 0.45D;
            if (item.IsClipped) confidence -= 0.2D;
            if (page.RotationDegrees != 0) confidence -= 0.03D;
            var evidence = new List<PdfInferenceEvidence> {
                new PdfInferenceEvidence(
                    item.HasGeometry ? "reading-order.visible-geometry" : "reading-order.source-sequence-fallback",
                    item.HasGeometry ? "The item was ordered from visible page geometry." : "The item has no placement geometry and retains source sequence after positioned content.",
                    item.HasGeometry ? 0.8D : -0.2D),
                new PdfInferenceEvidence(
                    item.SpansColumns ? "reading-order.spanning-band" : "reading-order.column-band",
                    item.SpansColumns ? "The item spans the page column band." : "The item was assigned to column " + item.ColumnIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".",
                    item.SpansColumns ? 0.7D : 0.6D)
            };
            if (page.RotationDegrees != 0) {
                evidence.Add(new PdfInferenceEvidence("reading-order.page-rotation", "Geometry was normalized through the page's " + page.RotationDegrees.ToString(System.Globalization.CultureInfo.InvariantCulture) + " degree rotation.", 0.7D));
            }
            if (item.IsClipped) {
                evidence.Add(new PdfInferenceEvidence("reading-order.crop-clipped", "Source geometry was clipped to the visible crop boundary.", -0.4D));
            }
            result[index] = new PdfLogicalReadingOrderItem(
                item.Kind, item.SourceIndex, item.PlacementIndex, index, item.ColumnIndex, item.SpansColumns,
                item.HasGeometry, item.IsClipped, item.Left, item.Top, item.Right, item.Bottom,
                confidence, evidence.AsReadOnly());
        }
        return ApplyCanonicalOrder(page, result);

        void AddBand(IEnumerable<Candidate> source, List<Candidate> destination, HashSet<Candidate> seen) {
            Candidate[] band = source.OrderBy(static item => item.Left).ThenBy(static item => item.Top).ToArray();
            if (band.Length == 0) return;
            double[] anchors = FindRepeatedColumnAnchors(
                band,
                pageWidth,
                consumeWork,
                cancellationCheck);
            if (anchors.Length < 2) {
                foreach (Candidate item in band.OrderBy(static item => item.Top).ThenBy(static item => item.Left).ThenBy(static item => item.Sequence)) {
                    item.ColumnIndex = 0;
                    if (seen.Add(item)) destination.Add(item);
                }
                return;
            }
            var columns = anchors.Select(static _ => new List<Candidate>()).ToArray();
            for (int itemIndex = 0; itemIndex < band.Length; itemIndex++) {
                Candidate item = band[itemIndex];
                int nearestIndex = 0;
                double nearestDistance = double.MaxValue;
                for (int anchorIndex = 0; anchorIndex < anchors.Length; anchorIndex++) {
                    double distance = Math.Abs(anchors[anchorIndex] - item.Left);
                    if (distance < nearestDistance) {
                        nearestDistance = distance;
                        nearestIndex = anchorIndex;
                    }
                }
                columns[nearestIndex].Add(item);
            }
            for (int columnIndex = 0; columnIndex < columns.Length; columnIndex++) {
                foreach (Candidate item in columns[columnIndex].OrderBy(static item => item.Top).ThenBy(static item => item.Left).ThenBy(static item => item.Sequence)) {
                    item.ColumnIndex = columnIndex;
                    if (seen.Add(item)) destination.Add(item);
                }
            }
        }
    }

    private static IReadOnlyList<PdfLogicalReadingOrderItem> ApplyCanonicalOrder(
        PdfLogicalPage page,
        PdfLogicalReadingOrderItem[] items) {
        if (page.RotationDegrees != 0 ||
            page.Analysis.ReadingOrder.Count == 0 ||
            items.Length < 2) return items;

        Dictionary<(long BaselineBucket, long XBucket, string Text), IReadOnlyList<CanonicalLinePosition>> canonicalLines =
            IndexCanonicalLines(page.Analysis.ReadingOrder);
        var ranked = new CanonicalRank[items.Length];
        var matchedIndexes = new List<int>();
        for (int index = 0; index < items.Length; index++) {
            bool matched = TryGetCanonicalPosition(page, items[index], canonicalLines, out long position);
            ranked[index] = new CanonicalRank(items[index], index, matched, matched ? position : 0D);
            if (matched) matchedIndexes.Add(index);
        }
        if (matchedIndexes.Count == 0) return items;

        for (int index = 0; index < ranked.Length; index++) {
            if (ranked[index].Matched) continue;
            int previous = -1;
            int next = -1;
            for (int matchIndex = 0; matchIndex < matchedIndexes.Count; matchIndex++) {
                int candidate = matchedIndexes[matchIndex];
                if (candidate < index) previous = candidate;
                else if (candidate > index) { next = candidate; break; }
            }

            double position;
            if (previous >= 0 && next >= 0 && ranked[previous].Position < ranked[next].Position) {
                double fraction = (double)(index - previous) / (next - previous);
                position = ranked[previous].Position + ((ranked[next].Position - ranked[previous].Position) * fraction);
            } else if (previous >= 0) {
                position = ranked[previous].Position + 50_000D + (index - previous);
            } else if (next >= 0) {
                position = ranked[next].Position - 50_000D - (next - index);
            } else {
                position = items[index].OrderIndex;
            }
            ranked[index] = ranked[index].WithPosition(position);
        }

        CanonicalRank[] ordered = ranked
            .OrderBy(static value => value.Position)
            .ThenBy(static value => value.OriginalIndex)
            .ToArray();
        var result = new PdfLogicalReadingOrderItem[ordered.Length];
        for (int index = 0; index < ordered.Length; index++) {
            PdfLogicalReadingOrderItem item = ordered[index].Item;
            IReadOnlyList<PdfInferenceEvidence> evidence = item.Evidence;
            if (ordered[index].Matched) {
                evidence = item.Evidence.Concat(new[] {
                    new PdfInferenceEvidence(
                        "reading-order.canonical-understanding",
                        "The item's position is owned by the canonical PDF understanding pipeline.",
                        0.9D)
                }).ToArray();
            }
            result[index] = new PdfLogicalReadingOrderItem(
                item.Kind,
                item.SourceIndex,
                item.PlacementIndex,
                index,
                item.ColumnIndex,
                item.SpansColumns,
                item.HasGeometry,
                item.IsClipped,
                item.Left,
                item.Top,
                item.Right,
                item.Bottom,
                item.Confidence,
                evidence);
        }
        return Array.AsReadOnly(result);
    }

    private static bool TryGetCanonicalPosition(
        PdfLogicalPage page,
        PdfLogicalReadingOrderItem item,
        Dictionary<(long BaselineBucket, long XBucket, string Text), IReadOnlyList<CanonicalLinePosition>> canonicalLines,
        out long position) {
        position = 0L;
        IReadOnlyList<PdfLogicalTextBlock>? lines = item.Kind switch {
            PdfLogicalReadingOrderKind.TextBlock => new[] { page.TextBlocks[item.SourceIndex] },
            PdfLogicalReadingOrderKind.Heading => new[] { page.Headings[item.SourceIndex].Line },
            PdfLogicalReadingOrderKind.Paragraph => page.Paragraphs[item.SourceIndex].Lines,
            PdfLogicalReadingOrderKind.ListItem => page.ListItems[item.SourceIndex].Lines,
            _ => null
        };
        if (lines is null || lines.Count == 0) return false;

        long best = long.MaxValue;
        for (int logicalLineIndex = 0; logicalLineIndex < lines.Count; logicalLineIndex++) {
            PdfLogicalTextBlock block = lines[logicalLineIndex];
            string text = PdfTextSimilarity.NormalizeSignature(block.Text);
            long baselineBucket = GetCanonicalBucket(block.BaselineY, 0.25D);
            long xBucket = GetCanonicalBucket(block.XStart, 0.5D);
            for (long baselineOffset = -1; baselineOffset <= 1; baselineOffset++) {
                for (long xOffset = -1; xOffset <= 1; xOffset++) {
                    if (!canonicalLines.TryGetValue(
                            (baselineBucket + baselineOffset, xBucket + xOffset, text),
                            out IReadOnlyList<CanonicalLinePosition>? candidates)) continue;
                    for (int candidateIndex = 0; candidateIndex < candidates.Count; candidateIndex++) {
                        CanonicalLinePosition candidate = candidates[candidateIndex];
                        if (Math.Abs(candidate.BaselineY - block.BaselineY) > 0.25D ||
                            Math.Abs(candidate.XStart - block.XStart) > 0.5D) continue;
                        best = Math.Min(best, candidate.Position);
                    }
                }
            }
        }
        if (best == long.MaxValue) return false;
        position = best;
        return true;
    }

    private static Dictionary<(long BaselineBucket, long XBucket, string Text), IReadOnlyList<CanonicalLinePosition>>
        IndexCanonicalLines(IReadOnlyList<PdfUnderstandingRegion> regions) {
        var index = new Dictionary<(long BaselineBucket, long XBucket, string Text), List<CanonicalLinePosition>>();
        for (int regionIndex = 0; regionIndex < regions.Count; regionIndex++) {
            IReadOnlyList<PdfUnderstandingLine> lines = regions[regionIndex].Lines;
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                PdfUnderstandingLine line = lines[lineIndex];
                var key = (
                    GetCanonicalBucket(line.BaselineY, 0.25D),
                    GetCanonicalBucket(line.XStart, 0.5D),
                    PdfTextSimilarity.NormalizeSignature(line.Text));
                if (!index.TryGetValue(key, out List<CanonicalLinePosition>? positions)) {
                    positions = new List<CanonicalLinePosition>();
                    index.Add(key, positions);
                }
                positions.Add(new CanonicalLinePosition(
                    line.BaselineY,
                    line.XStart,
                    checked((long)regionIndex * 100_000L + lineIndex)));
            }
        }
        return index.ToDictionary(
            static pair => pair.Key,
            static pair => (IReadOnlyList<CanonicalLinePosition>)pair.Value.AsReadOnly());
    }

    private static long GetCanonicalBucket(double value, double width) {
        if (double.IsNaN(value)) return 0L;
        double bucket = Math.Floor(value / width);
        if (bucket <= long.MinValue) return long.MinValue + 1;
        if (bucket >= long.MaxValue) return long.MaxValue - 1;
        return (long)bucket;
    }

    private readonly struct CanonicalLinePosition {
        internal CanonicalLinePosition(double baselineY, double xStart, long position) {
            BaselineY = baselineY;
            XStart = xStart;
            Position = position;
        }

        internal double BaselineY { get; }
        internal double XStart { get; }
        internal long Position { get; }
    }

    private readonly struct CanonicalRank {
        internal CanonicalRank(PdfLogicalReadingOrderItem item, int originalIndex, bool matched, double position) {
            Item = item;
            OriginalIndex = originalIndex;
            Matched = matched;
            Position = position;
        }

        internal PdfLogicalReadingOrderItem Item { get; }
        internal int OriginalIndex { get; }
        internal bool Matched { get; }
        internal double Position { get; }
        internal CanonicalRank WithPosition(double position) => new CanonicalRank(Item, OriginalIndex, Matched, position);
    }

    private static List<Candidate> BuildCandidates(PdfLogicalPage page, PdfLogicalReadingOrderScope scope) {
        var result = new List<Candidate>();
        var representedTextBlocks = new HashSet<PdfLogicalTextBlock>();
        for (int index = 0; index < page.Headings.Count; index++) representedTextBlocks.Add(page.Headings[index].Line);
        for (int index = 0; index < page.Paragraphs.Count; index++) {
            foreach (PdfLogicalTextBlock line in page.Paragraphs[index].Lines) representedTextBlocks.Add(line);
        }
        for (int index = 0; index < page.ListItems.Count; index++) {
            foreach (PdfLogicalTextBlock line in page.ListItems[index].Lines) representedTextBlocks.Add(line);
        }
        for (int blockIndex = 0; blockIndex < page.TextBlocks.Count; blockIndex++) {
            PdfLogicalTextBlock block = page.TextBlocks[blockIndex];
            if (page.Tables.Any(table => IsOwnedByTable(page, block, table))) representedTextBlocks.Add(block);
        }
        int sequence = 0;
        for (int index = 0; index < page.TextBlocks.Count; index++) {
            PdfLogicalTextBlock block = page.TextBlocks[index];
            if (scope == PdfLogicalReadingOrderScope.SemanticBody &&
                block.Kind is (PdfLogicalElementKind.Header or PdfLogicalElementKind.Footer)) continue;
            if (!representedTextBlocks.Contains(block)) AddText(PdfLogicalReadingOrderKind.TextBlock, index, new[] { block });
        }
        for (int index = 0; index < page.Headings.Count; index++) AddText(PdfLogicalReadingOrderKind.Heading, index, new[] { page.Headings[index].Line });
        for (int index = 0; index < page.Paragraphs.Count; index++) AddText(PdfLogicalReadingOrderKind.Paragraph, index, page.Paragraphs[index].Lines);
        for (int index = 0; index < page.ListItems.Count; index++) AddText(PdfLogicalReadingOrderKind.ListItem, index, page.ListItems[index].Lines);
        for (int index = 0; index < page.Tables.Count; index++) {
            PdfLogicalTable table = page.Tables[index];
            if (table.VisualBounds is PdfLogicalVisualBounds visualBounds) {
                AddVisual(PdfLogicalReadingOrderKind.Table, index, -1, visualBounds.Left, visualBounds.Top, visualBounds.Right, visualBounds.Bottom);
                continue;
            }
            double left = table.Columns.Count == 0 ? 0D : table.Columns.Min(static column => Math.Min(column.From, column.To));
            double right = table.Columns.Count == 0 ? 0D : table.Columns.Max(static column => Math.Max(column.From, column.To));
            Add(PdfLogicalReadingOrderKind.Table, index, -1, left, Math.Min(table.YTop, table.YBottom), right, Math.Max(table.YTop, table.YBottom));
        }
        for (int index = 0; index < page.Images.Count; index++) {
            PdfLogicalImage image = page.Images[index];
            if (image.Placements.Count == 0) AddMissing(PdfLogicalReadingOrderKind.Image, index, -1);
            for (int placementIndex = 0; placementIndex < image.Placements.Count; placementIndex++) {
                PdfImagePlacement placement = image.Placements[placementIndex];
                Add(PdfLogicalReadingOrderKind.Image, index, placementIndex, placement.X, placement.Y, placement.X + placement.Width, placement.Y + placement.Height);
            }
        }
        for (int index = 0; index < page.Links.Count; index++) {
            PdfLogicalLinkAnnotation link = page.Links[index];
            Add(PdfLogicalReadingOrderKind.Link, index, -1, link.X1, link.Y1, link.X2, link.Y2);
        }
        for (int index = 0; index < page.FormWidgets.Count; index++) {
            PdfLogicalFormWidget widget = page.FormWidgets[index];
            Add(PdfLogicalReadingOrderKind.FormWidget, index, -1, widget.X1, widget.Y1, widget.X2, widget.Y2);
        }
        return result;

        void AddText(PdfLogicalReadingOrderKind kind, int sourceIndex, IReadOnlyList<PdfLogicalTextBlock> lines) {
            if (lines.Count == 0) { AddMissing(kind, sourceIndex, -1); return; }
            PdfLogicalVisualBounds[] directBounds = lines.Select(static line => line.VisualBounds).Where(static bounds => bounds is not null).Cast<PdfLogicalVisualBounds>().ToArray();
            if (directBounds.Length == lines.Count) {
                AddVisual(
                    kind,
                    sourceIndex,
                    -1,
                    directBounds.Min(static bounds => bounds.Left),
                    directBounds.Min(static bounds => bounds.Top),
                    directBounds.Max(static bounds => bounds.Right),
                    directBounds.Max(static bounds => bounds.Bottom));
                return;
            }
            double left = double.MaxValue;
            double right = double.MinValue;
            double bottom = double.MaxValue;
            double top = double.MinValue;
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                PdfLogicalTextBlock line = lines[lineIndex];
                left = Math.Min(left, line.XStart);
                right = Math.Max(right, line.XEnd);
                bottom = Math.Min(bottom, line.BaselineY - Math.Max(1D, line.FontSize * 0.25D));
                top = Math.Max(top, line.BaselineY + Math.Max(1D, line.FontSize));
            }
            Add(kind, sourceIndex, -1, left, bottom, right, top);
        }

        void AddVisual(PdfLogicalReadingOrderKind kind, int sourceIndex, int placementIndex, double left, double top, double right, double bottom) {
            (double width, double height) = page.GetVisualPageSize();
            double clippedLeft = Math.Max(0D, Math.Min(width, left));
            double clippedTop = Math.Max(0D, Math.Min(height, top));
            double clippedRight = Math.Max(0D, Math.Min(width, right));
            double clippedBottom = Math.Max(0D, Math.Min(height, bottom));
            if (clippedRight <= clippedLeft || clippedBottom <= clippedTop) {
                AddMissing(kind, sourceIndex, placementIndex);
                return;
            }
            bool clipped = Math.Abs(clippedLeft - left) > 0.001D || Math.Abs(clippedTop - top) > 0.001D || Math.Abs(clippedRight - right) > 0.001D || Math.Abs(clippedBottom - bottom) > 0.001D;
            result.Add(new Candidate(kind, sourceIndex, placementIndex, sequence++, clippedLeft, clippedTop, clippedRight, clippedBottom, hasGeometry: true, clipped));
        }

        void Add(PdfLogicalReadingOrderKind kind, int sourceIndex, int placementIndex, double left, double bottom, double right, double top) {
            if (!IsFinite(left) || !IsFinite(bottom) || !IsFinite(right) || !IsFinite(top) || right <= left || top <= bottom) {
                AddMissing(kind, sourceIndex, placementIndex);
                return;
            }
            PdfVisualBounds visual = page.TransformBoundsToVisual(left, bottom, right, top);
            (double width, double height) = page.GetVisualPageSize();
            double clippedLeft = Math.Max(0D, Math.Min(width, visual.Left));
            double clippedTop = Math.Max(0D, Math.Min(height, visual.Top));
            double clippedRight = Math.Max(0D, Math.Min(width, visual.Right));
            double clippedBottom = Math.Max(0D, Math.Min(height, visual.Bottom));
            bool clipped = Math.Abs(clippedLeft - visual.Left) > 0.001D || Math.Abs(clippedTop - visual.Top) > 0.001D || Math.Abs(clippedRight - visual.Right) > 0.001D || Math.Abs(clippedBottom - visual.Bottom) > 0.001D;
            if (clippedRight <= clippedLeft || clippedBottom <= clippedTop) {
                AddMissing(kind, sourceIndex, placementIndex);
                return;
            }
            result.Add(new Candidate(kind, sourceIndex, placementIndex, sequence++, clippedLeft, clippedTop, clippedRight, clippedBottom, hasGeometry: true, clipped));
        }

        void AddMissing(PdfLogicalReadingOrderKind kind, int sourceIndex, int placementIndex) =>
            result.Add(new Candidate(kind, sourceIndex, placementIndex, sequence++, 0D, 0D, 0D, 0D, hasGeometry: false, isClipped: false));
    }

    private static bool IsOwnedByTable(PdfLogicalPage page, PdfLogicalTextBlock block, PdfLogicalTable table) {
        if (!TryGetVisualBounds(page, block, out PdfVisualBounds blockBounds) ||
            !TryGetVisualBounds(page, table, out PdfVisualBounds tableBounds)) return false;
        double blockWidth = blockBounds.Right - blockBounds.Left;
        if (blockWidth <= 0.001D) return false;
        double horizontalOverlap = Math.Max(0D, Math.Min(blockBounds.Right, tableBounds.Right) - Math.Max(blockBounds.Left, tableBounds.Left));
        if (horizontalOverlap + 0.001D < blockWidth * 0.5D) return false;
        double verticalPadding = Math.Max(1D, block.FontSize);
        double blockCenter = (blockBounds.Top + blockBounds.Bottom) / 2D;
        return blockCenter >= tableBounds.Top - verticalPadding && blockCenter <= tableBounds.Bottom + verticalPadding;
    }

    private static bool TryGetVisualBounds(PdfLogicalPage page, PdfLogicalTextBlock block, out PdfVisualBounds bounds) {
        if (block.VisualBounds is PdfLogicalVisualBounds visual) {
            bounds = new PdfVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
            return visual.Right > visual.Left && visual.Bottom > visual.Top;
        }
        if (block.XEnd <= block.XStart) { bounds = default; return false; }
        double bottom = block.BaselineY - Math.Max(1D, block.FontSize * 0.25D);
        double top = block.BaselineY + Math.Max(1D, block.FontSize);
        bounds = page.TransformBoundsToVisual(block.XStart, bottom, block.XEnd, top);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static bool TryGetVisualBounds(PdfLogicalPage page, PdfLogicalTable table, out PdfVisualBounds bounds) {
        if (table.VisualBounds is PdfLogicalVisualBounds visual) {
            bounds = new PdfVisualBounds(visual.Left, visual.Top, visual.Right, visual.Bottom);
            return visual.Right > visual.Left && visual.Bottom > visual.Top;
        }
        if (table.Columns.Count == 0) { bounds = default; return false; }
        double left = table.Columns.Min(static column => Math.Min(column.From, column.To));
        double right = table.Columns.Max(static column => Math.Max(column.From, column.To));
        double bottom = Math.Min(table.YBottom, table.YTop);
        double top = Math.Max(table.YBottom, table.YTop);
        if (right <= left || top <= bottom) { bounds = default; return false; }
        bounds = page.TransformBoundsToVisual(left, bottom, right, top);
        return bounds.Right > bounds.Left && bounds.Bottom > bounds.Top;
    }

    private static double[] FindRepeatedColumnAnchors(
        IEnumerable<Candidate> source,
        double pageWidth,
        Action<long>? consumeWork,
        Action? cancellationCheck) {
        cancellationCheck?.Invoke();
        Candidate[] candidates = source.OrderBy(static item => item.Left).ToArray();
        if (candidates.Length > 0) consumeWork?.Invoke(candidates.Length);
        if (candidates.Length < 4) return Array.Empty<double>();
        double[] widths = candidates.Select(static item => Math.Max(1D, item.Right - item.Left)).OrderBy(static width => width).ToArray();
        double tolerance = Math.Max(18D, Math.Min(pageWidth * 0.12D, widths[widths.Length / 2] * 0.25D));
        var clusters = new List<ColumnAnchorCluster>();
        for (int index = 0; index < candidates.Length; index++) {
            cancellationCheck?.Invoke();
            Candidate item = candidates[index];
            ColumnAnchorCluster? cluster = clusters.Count == 0 ? null : clusters[clusters.Count - 1];
            if (cluster is null || Math.Abs(cluster.Centroid - item.Left) > tolerance) {
                cluster = new ColumnAnchorCluster();
                clusters.Add(cluster);
            }
            cluster.Add(item.Left);
        }
        double[] repeated = clusters.Where(static cluster => cluster.Count >= 2)
            .Select(static cluster => cluster.Centroid)
            .OrderBy(static left => left)
            .ToArray();
        double minimumSeparation = Math.Max(72D, pageWidth * 0.2D);
        var columnRegions = new List<List<double>>();
        for (int index = 0; index < repeated.Length; index++) {
            double anchor = repeated[index];
            if (columnRegions.Count == 0 || anchor - columnRegions[columnRegions.Count - 1][columnRegions[columnRegions.Count - 1].Count - 1] >= minimumSeparation) {
                columnRegions.Add(new List<double> { anchor });
            } else {
                columnRegions[columnRegions.Count - 1].Add(anchor);
            }
        }
        if (columnRegions.Count < 2) return Array.Empty<double>();
        return columnRegions.Select(static region => region.Min()).ToArray();
    }

    private static bool IsCenteredBandDivider(Candidate item, double[] anchors, double pageWidth) {
        if (anchors.Length < 2 || item.Kind is not (PdfLogicalReadingOrderKind.Heading or PdfLogicalReadingOrderKind.Paragraph or PdfLogicalReadingOrderKind.TextBlock)) return false;
        double center = (item.Left + item.Right) / 2D;
        bool pageCentered = Math.Abs(center - (pageWidth / 2D)) <= pageWidth * 0.12D;
        if (item.Kind == PdfLogicalReadingOrderKind.Heading) return pageCentered;
        double nearestAnchor = anchors.Min(anchor => Math.Abs(anchor - item.Left));
        return pageCentered && nearestAnchor > Math.Max(24D, pageWidth * 0.1D);
    }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private sealed class ColumnAnchorCluster {
        private double _leftSum;

        internal int Count { get; private set; }
        internal double Centroid => Count == 0 ? 0D : _leftSum / Count;

        internal void Add(double left) {
            _leftSum += left;
            Count++;
        }
    }

    private sealed class Candidate {
        internal Candidate(PdfLogicalReadingOrderKind kind, int sourceIndex, int placementIndex, int sequence, double left, double top, double right, double bottom, bool hasGeometry, bool isClipped) {
            Kind = kind; SourceIndex = sourceIndex; PlacementIndex = placementIndex; Sequence = sequence; Left = left; Top = top; Right = right; Bottom = bottom; HasGeometry = hasGeometry; IsClipped = isClipped;
        }
        internal PdfLogicalReadingOrderKind Kind { get; }
        internal int SourceIndex { get; }
        internal int PlacementIndex { get; }
        internal int Sequence { get; }
        internal double Left { get; }
        internal double Top { get; }
        internal double Right { get; }
        internal double Bottom { get; }
        internal bool HasGeometry { get; }
        internal bool IsClipped { get; }
        internal int ColumnIndex { get; set; }
        internal bool SpansColumns { get; set; }
    }
}
