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
    public static IReadOnlyList<PdfLogicalReadingOrderItem> Analyze(PdfLogicalPage page) {
        Guard.NotNull(page, nameof(page));
        var candidates = BuildCandidates(page);
        if (candidates.Count == 0) return Array.Empty<PdfLogicalReadingOrderItem>();

        (double pageWidth, double pageHeight) = page.GetVisualPageSize();
        var positioned = candidates.Where(static item => item.HasGeometry).ToArray();
        var unpositioned = candidates.Where(static item => !item.HasGeometry).OrderBy(static item => item.Sequence).ToArray();
        double[] pageColumnAnchors = FindRepeatedColumnAnchors(
            positioned.Where(item => item.Right - item.Left < Math.Max(1D, pageWidth) * SpanningWidthRatio),
            pageWidth);
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
        return Array.AsReadOnly(result);

        void AddBand(IEnumerable<Candidate> source, List<Candidate> destination, HashSet<Candidate> seen) {
            Candidate[] band = source.OrderBy(static item => item.Left).ThenBy(static item => item.Top).ToArray();
            if (band.Length == 0) return;
            double[] anchors = FindRepeatedColumnAnchors(band, pageWidth);
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

    private static List<Candidate> BuildCandidates(PdfLogicalPage page) {
        var result = new List<Candidate>();
        var semanticTextBlocks = new HashSet<PdfLogicalTextBlock>();
        for (int index = 0; index < page.Headings.Count; index++) semanticTextBlocks.Add(page.Headings[index].Line);
        for (int index = 0; index < page.Paragraphs.Count; index++) {
            foreach (PdfLogicalTextBlock line in page.Paragraphs[index].Lines) semanticTextBlocks.Add(line);
        }
        for (int index = 0; index < page.ListItems.Count; index++) semanticTextBlocks.Add(page.ListItems[index].Line);
        int sequence = 0;
        for (int index = 0; index < page.TextBlocks.Count; index++) {
            if (!semanticTextBlocks.Contains(page.TextBlocks[index])) AddText(PdfLogicalReadingOrderKind.TextBlock, index, new[] { page.TextBlocks[index] });
        }
        for (int index = 0; index < page.Headings.Count; index++) AddText(PdfLogicalReadingOrderKind.Heading, index, new[] { page.Headings[index].Line });
        for (int index = 0; index < page.Paragraphs.Count; index++) AddText(PdfLogicalReadingOrderKind.Paragraph, index, page.Paragraphs[index].Lines);
        for (int index = 0; index < page.ListItems.Count; index++) AddText(PdfLogicalReadingOrderKind.ListItem, index, new[] { page.ListItems[index].Line });
        for (int index = 0; index < page.Tables.Count; index++) {
            PdfLogicalTable table = page.Tables[index];
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

    private static double[] FindRepeatedColumnAnchors(IEnumerable<Candidate> source, double pageWidth) {
        Candidate[] candidates = source.OrderBy(static item => item.Left).ToArray();
        if (candidates.Length < 4) return Array.Empty<double>();
        double[] widths = candidates.Select(static item => Math.Max(1D, item.Right - item.Left)).OrderBy(static width => width).ToArray();
        double tolerance = Math.Max(18D, Math.Min(pageWidth * 0.12D, widths[widths.Length / 2] * 0.25D));
        var clusters = new List<List<Candidate>>();
        for (int index = 0; index < candidates.Length; index++) {
            Candidate item = candidates[index];
            List<Candidate>? cluster = clusters.FirstOrDefault(existing => Math.Abs(existing.Average(static candidate => candidate.Left) - item.Left) <= tolerance);
            if (cluster is null) { cluster = new List<Candidate>(); clusters.Add(cluster); }
            cluster.Add(item);
        }
        double[] repeated = clusters.Where(static cluster => cluster.Count >= 2)
            .Select(static cluster => cluster.Average(static item => item.Left))
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
