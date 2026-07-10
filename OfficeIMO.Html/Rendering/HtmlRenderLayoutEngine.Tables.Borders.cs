using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderBoxStyle CreateCollapsedCellPaintStyle(HtmlRenderBoxStyle source) {
        HtmlRenderBoxStyle result = source.Clone();
        result.Borders = result.Borders.WithUniformColor(OfficeColor.Transparent);
        return result;
    }

    private static void AddCollapsedTableBorders(
        ICollection<HtmlRenderVisual> visuals,
        IElement table,
        IReadOnlyList<TableRowLayout> rows,
        IReadOnlyList<double> columnWidths,
        IReadOnlyList<double> columnOffsets,
        double x,
        double y) {
        var winners = new Dictionary<CollapsedBorderKey, HtmlRenderBorderSide>();
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                for (int column = cell.Column; column < cell.Column + cell.Span && column < columnWidths.Count; column++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex, column), cell.Style.Borders.Top);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex + cell.RowSpan, column), cell.Style.Borders.Bottom);
                }
                for (int row = rowIndex; row < rowIndex + cell.RowSpan && row < rows.Count; row++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, cell.Column, row), cell.Style.Borders.Left);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, cell.Column + cell.Span, row), cell.Style.Borders.Right);
                }
            }
        }

        double[] rowOffsets = new double[rows.Count + 1];
        for (int row = 1; row < rowOffsets.Length; row++) rowOffsets[row] = rowOffsets[row - 1] + rows[row - 1].Height;
        string source = HtmlRenderStyleResolver.DescribeSource(table) + ":collapsed-border";
        foreach (KeyValuePair<CollapsedBorderKey, HtmlRenderBorderSide> entry in winners
            .OrderBy(pair => pair.Key.Horizontal ? 0 : 1)
            .ThenBy(pair => pair.Key.Boundary)
            .ThenBy(pair => pair.Key.Segment)) {
            HtmlRenderBorderSide border = entry.Value;
            if (!border.IsPainted) continue;
            CollapsedBorderKey key = entry.Key;
            if (key.Horizontal) {
                if (key.Boundary >= rowOffsets.Length || key.Segment >= columnWidths.Count) continue;
                AddCollapsedBorderLine(
                    visuals,
                    horizontal: true,
                    x + columnOffsets[key.Segment],
                    y + rowOffsets[key.Boundary],
                    columnWidths[key.Segment],
                    border,
                    source + "-h-" + key.Boundary + "-" + key.Segment);
            } else {
                if (key.Boundary > columnOffsets.Count || key.Segment >= rows.Count) continue;
                double boundaryX = key.Boundary == columnWidths.Count
                    ? x + columnWidths.Sum()
                    : x + columnOffsets[key.Boundary];
                AddCollapsedBorderLine(
                    visuals,
                    horizontal: false,
                    boundaryX,
                    y + rowOffsets[key.Segment],
                    rows[key.Segment].Height,
                    border,
                    source + "-v-" + key.Boundary + "-" + key.Segment);
            }
        }
    }

    private static void AddCollapsedBorderCandidate(
        IDictionary<CollapsedBorderKey, HtmlRenderBorderSide> winners,
        CollapsedBorderKey key,
        HtmlRenderBorderSide candidate) {
        if (!winners.TryGetValue(key, out HtmlRenderBorderSide current) || CompareCollapsedBorders(candidate, current) >= 0) winners[key] = candidate;
    }

    private static int CompareCollapsedBorders(HtmlRenderBorderSide left, HtmlRenderBorderSide right) {
        if (left.Style == "hidden" || right.Style == "hidden") {
            if (left.Style == right.Style) return 0;
            return left.Style == "hidden" ? 1 : -1;
        }
        int width = left.Width.CompareTo(right.Width);
        if (width != 0) return width;
        return CollapsedBorderStyleRank(left.Style).CompareTo(CollapsedBorderStyleRank(right.Style));
    }

    private static int CollapsedBorderStyleRank(string style) => style switch {
        "double" => 4,
        "solid" => 3,
        "dashed" => 2,
        "dotted" => 1,
        _ => 0
    };

    private static void AddCollapsedBorderLine(
        ICollection<HtmlRenderVisual> visuals,
        bool horizontal,
        double x,
        double y,
        double length,
        HtmlRenderBorderSide border,
        string source) {
        if (length <= 0.0001D || border.Width <= 0D) return;
        if (border.Style != "double") {
            AddCollapsedBorderStroke(visuals, horizontal, x, y, length, border.Color, border.Width, border.Style, source);
            return;
        }
        double strokeWidth = Math.Max(0.01D, border.Width / 3D);
        AddCollapsedBorderStroke(visuals, horizontal, x, y, length, border.Color, strokeWidth, "solid", source + "-outer");
        double inset = border.Width * 2D / 3D;
        AddCollapsedBorderStroke(visuals, horizontal, horizontal ? x : x + inset, horizontal ? y + inset : y, length, border.Color, strokeWidth, "solid", source + "-inner");
    }

    private static void AddCollapsedBorderStroke(
        ICollection<HtmlRenderVisual> visuals,
        bool horizontal,
        double x,
        double y,
        double length,
        OfficeColor color,
        double width,
        string style,
        string source) {
        OfficeShape shape = horizontal
            ? OfficeShape.Line(0D, 0D, length, 0.0001D)
            : OfficeShape.Line(0D, 0D, 0.0001D, length);
        shape.StrokeColor = color;
        shape.StrokeWidth = width;
        shape.StrokeDashStyle = MapStrokeDashStyle(style);
        shape.FillColor = null;
        visuals.Add(new HtmlRenderShape(shape, x, y, visuals.Count, source: source));
    }

    private readonly struct CollapsedBorderKey : IEquatable<CollapsedBorderKey> {
        internal CollapsedBorderKey(bool horizontal, int boundary, int segment) {
            Horizontal = horizontal;
            Boundary = boundary;
            Segment = segment;
        }

        internal bool Horizontal { get; }
        internal int Boundary { get; }
        internal int Segment { get; }

        public bool Equals(CollapsedBorderKey other) => Horizontal == other.Horizontal && Boundary == other.Boundary && Segment == other.Segment;
        public override bool Equals(object? obj) => obj is CollapsedBorderKey other && Equals(other);
        public override int GetHashCode() {
            unchecked {
                int hash = Horizontal ? 1 : 0;
                hash = (hash * 397) ^ Boundary;
                hash = (hash * 397) ^ Segment;
                return hash;
            }
        }
    }
}
