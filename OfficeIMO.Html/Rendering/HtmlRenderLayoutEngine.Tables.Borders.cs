using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static HtmlRenderBoxStyle CreateCollapsedCellPaintStyle(HtmlRenderBoxStyle source) {
        HtmlRenderBoxStyle result = source.Clone();
        result.Borders = result.Borders.WithUniformColor(OfficeColor.Transparent);
        return result;
    }

    private void AddCollapsedTableBorders(
        ICollection<HtmlRenderVisual> visuals,
        IElement table,
        HtmlRenderBoxStyle tableStyle,
        IReadOnlyList<TableRowLayout> rows,
        IReadOnlyList<double> columnWidths,
        IReadOnlyList<double> columnOffsets,
        double x,
        double y) {
        var winners = new Dictionary<CollapsedBorderKey, CollapsedBorderCandidate>();
        int sourceOrder = 0;
        AddHorizontalBorderRange(winners, 0, 0, columnWidths.Count, tableStyle.Borders.Top, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddHorizontalBorderRange(winners, rows.Count, 0, columnWidths.Count, tableStyle.Borders.Bottom, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddVerticalBorderRange(winners, 0, 0, rows.Count, tableStyle.Borders.Left, CollapsedBorderOrigin.Table, ref sourceOrder);
        AddVerticalBorderRange(winners, columnWidths.Count, 0, rows.Count, tableStyle.Borders.Right, CollapsedBorderOrigin.Table, ref sourceOrder);

        AddCollapsedColumnBorders(winners, table, tableStyle, rows.Count, columnWidths.Count, columnWidths.Sum(), ref sourceOrder);
        AddCollapsedRowGroupBorders(winners, rows, columnWidths.Count, ref sourceOrder);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            HtmlRenderBorderEdges rowBorders = rows[rowIndex].Style.Borders;
            AddHorizontalBorderRange(winners, rowIndex, 0, columnWidths.Count, rowBorders.Top, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddHorizontalBorderRange(winners, rowIndex + 1, 0, columnWidths.Count, rowBorders.Bottom, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddVerticalBorderRange(winners, 0, rowIndex, rowIndex + 1, rowBorders.Left, CollapsedBorderOrigin.Row, ref sourceOrder);
            AddVerticalBorderRange(winners, columnWidths.Count, rowIndex, rowIndex + 1, rowBorders.Right, CollapsedBorderOrigin.Row, ref sourceOrder);
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                for (int column = cell.Column; column < cell.Column + cell.Span && column < columnWidths.Count; column++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex, column), cell.Style.Borders.Top, CollapsedBorderOrigin.Cell, sourceOrder++);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, rowIndex + cell.RowSpan, column), cell.Style.Borders.Bottom, CollapsedBorderOrigin.Cell, sourceOrder++);
                }
                for (int row = rowIndex; row < rowIndex + cell.RowSpan && row < rows.Count; row++) {
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, cell.Column, row), cell.Style.Borders.Left, CollapsedBorderOrigin.Cell, sourceOrder++);
                    AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, cell.Column + cell.Span, row), cell.Style.Borders.Right, CollapsedBorderOrigin.Cell, sourceOrder++);
                }
            }
        }

        double[] rowOffsets = new double[rows.Count + 1];
        for (int row = 1; row < rowOffsets.Length; row++) rowOffsets[row] = rowOffsets[row - 1] + rows[row - 1].Height;
        string source = HtmlRenderStyleResolver.DescribeSource(table) + ":collapsed-border";
        foreach (KeyValuePair<CollapsedBorderKey, CollapsedBorderCandidate> entry in winners
            .OrderBy(pair => pair.Key.Horizontal ? 0 : 1)
            .ThenBy(pair => pair.Key.Boundary)
            .ThenBy(pair => pair.Key.Segment)) {
            HtmlRenderBorderSide border = entry.Value.Border;
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

    private void AddCollapsedColumnBorders(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        IElement table,
        HtmlRenderBoxStyle tableStyle,
        int rowCount,
        int columnCount,
        double contentWidth,
        ref int sourceOrder) {
        var groupStyles = new Dictionary<IElement, HtmlRenderBoxStyle>();
        var groupRanges = new Dictionary<IElement, CollapsedColumnGroupRange>();
        int column = 0;
        foreach (IElement element in table.QuerySelectorAll("col").Where(candidate => BelongsToTableColumn(candidate, table))) {
            int span = Math.Min(ReadSpan(element.GetAttribute("span"), columnCount), columnCount - column);
            if (span <= 0) break;
            IElement? group = FindColumnGroup(element, table);
            HtmlRenderBoxStyle parentStyle = tableStyle;
            if (group != null) {
                if (!groupStyles.TryGetValue(group, out HtmlRenderBoxStyle? groupStyle)) {
                    groupStyle = _styleResolver.Resolve(group, contentWidth, tableStyle);
                    groupStyles[group] = groupStyle;
                    _layoutStyles[group] = groupStyle.Clone();
                }
                parentStyle = groupStyle;
                if (groupRanges.TryGetValue(group, out CollapsedColumnGroupRange range)) {
                    groupRanges[group] = new CollapsedColumnGroupRange(range.Start, column + span, groupStyle);
                } else {
                    groupRanges[group] = new CollapsedColumnGroupRange(column, column + span, groupStyle);
                }
            }

            HtmlRenderBoxStyle columnStyle = _styleResolver.Resolve(element, contentWidth, parentStyle);
            _layoutStyles[element] = columnStyle.Clone();
            AddCollapsedColumnRange(winners, column, column + span, rowCount, columnStyle.Borders, CollapsedBorderOrigin.Column, ref sourceOrder);
            column += span;
            if (column >= columnCount) break;
        }

        foreach (CollapsedColumnGroupRange range in groupRanges.Values.OrderBy(range => range.Start)) {
            AddCollapsedColumnRange(winners, range.Start, range.End, rowCount, range.Style.Borders, CollapsedBorderOrigin.ColumnGroup, ref sourceOrder);
        }
    }

    private static IElement? FindColumnGroup(IElement column, IElement table) {
        IElement? current = column.ParentElement;
        while (current != null && !ReferenceEquals(current, table)) {
            if (string.Equals(current.TagName, "colgroup", StringComparison.OrdinalIgnoreCase)) return current;
            current = current.ParentElement;
        }
        return null;
    }

    private void AddCollapsedColumnRange(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        int start,
        int end,
        int rowCount,
        HtmlRenderBorderEdges borders,
        CollapsedBorderOrigin origin,
        ref int sourceOrder) {
        AddHorizontalBorderRange(winners, 0, start, end, borders.Top, origin, ref sourceOrder);
        AddHorizontalBorderRange(winners, rowCount, start, end, borders.Bottom, origin, ref sourceOrder);
        AddVerticalBorderRange(winners, start, 0, rowCount, borders.Left, origin, ref sourceOrder);
        AddVerticalBorderRange(winners, end, 0, rowCount, borders.Right, origin, ref sourceOrder);
    }

    private void AddCollapsedRowGroupBorders(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        IReadOnlyList<TableRowLayout> rows,
        int columnCount,
        ref int sourceOrder) {
        int start = 0;
        while (start < rows.Count) {
            IElement? group = rows[start].GroupElement;
            if (group == null || rows[start].GroupStyle == null) {
                start++;
                continue;
            }
            int end = start + 1;
            while (end < rows.Count && ReferenceEquals(rows[end].GroupElement, group)) end++;
            HtmlRenderBorderEdges borders = rows[start].GroupStyle!.Borders;
            AddHorizontalBorderRange(winners, start, 0, columnCount, borders.Top, CollapsedBorderOrigin.RowGroup, ref sourceOrder);
            AddHorizontalBorderRange(winners, end, 0, columnCount, borders.Bottom, CollapsedBorderOrigin.RowGroup, ref sourceOrder);
            AddVerticalBorderRange(winners, 0, start, end, borders.Left, CollapsedBorderOrigin.RowGroup, ref sourceOrder);
            AddVerticalBorderRange(winners, columnCount, start, end, borders.Right, CollapsedBorderOrigin.RowGroup, ref sourceOrder);
            start = end;
        }
    }

    private void AddHorizontalBorderRange(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        int boundary,
        int start,
        int end,
        HtmlRenderBorderSide border,
        CollapsedBorderOrigin origin,
        ref int sourceOrder) {
        for (int segment = start; segment < end; segment++) {
            AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(true, boundary, segment), border, origin, sourceOrder++);
        }
    }

    private void AddVerticalBorderRange(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        int boundary,
        int start,
        int end,
        HtmlRenderBorderSide border,
        CollapsedBorderOrigin origin,
        ref int sourceOrder) {
        for (int segment = start; segment < end; segment++) {
            AddCollapsedBorderCandidate(winners, new CollapsedBorderKey(false, boundary, segment), border, origin, sourceOrder++);
        }
    }

    private void AddCollapsedBorderCandidate(
        IDictionary<CollapsedBorderKey, CollapsedBorderCandidate> winners,
        CollapsedBorderKey key,
        HtmlRenderBorderSide border,
        CollapsedBorderOrigin origin,
        int sourceOrder) {
        if (border.Style == "none") return;
        var candidate = new CollapsedBorderCandidate(border, origin, sourceOrder);
        if (winners.TryGetValue(key, out CollapsedBorderCandidate current)) {
            if (CompareCollapsedBorders(candidate, current) >= 0) winners[key] = candidate;
            return;
        }
        if (winners.Count >= _options.MaxCollapsedTableBorderSegments) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.CollapsedTableBorderLimitExceeded,
                "Collapsed table border resolution exceeded the configured segment limit.",
                nameof(HtmlRenderOptions.MaxCollapsedTableBorderSegments),
                winners.Count + 1L,
                _options.MaxCollapsedTableBorderSegments);
        }
        winners[key] = candidate;
    }

    private static int CompareCollapsedBorders(CollapsedBorderCandidate left, CollapsedBorderCandidate right) {
        if (left.Border.Style == "hidden" || right.Border.Style == "hidden") {
            if (left.Border.Style != right.Border.Style) return left.Border.Style == "hidden" ? 1 : -1;
        }
        if (left.Border.Style == "none" || right.Border.Style == "none") {
            if (left.Border.Style != right.Border.Style) return left.Border.Style == "none" ? -1 : 1;
        }
        int width = left.Border.Width.CompareTo(right.Border.Width);
        if (width != 0) return width;
        int style = CollapsedBorderStyleRank(left.Border.Style).CompareTo(CollapsedBorderStyleRank(right.Border.Style));
        if (style != 0) return style;
        int origin = left.Origin.CompareTo(right.Origin);
        return origin != 0 ? origin : left.SourceOrder.CompareTo(right.SourceOrder);
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

    private enum CollapsedBorderOrigin {
        Table = 1,
        ColumnGroup = 2,
        Column = 3,
        RowGroup = 4,
        Row = 5,
        Cell = 6
    }

    private readonly struct CollapsedBorderCandidate {
        internal CollapsedBorderCandidate(HtmlRenderBorderSide border, CollapsedBorderOrigin origin, int sourceOrder) {
            Border = border;
            Origin = origin;
            SourceOrder = sourceOrder;
        }

        internal HtmlRenderBorderSide Border { get; }
        internal CollapsedBorderOrigin Origin { get; }
        internal int SourceOrder { get; }
    }

    private readonly struct CollapsedColumnGroupRange {
        internal CollapsedColumnGroupRange(int start, int end, HtmlRenderBoxStyle style) {
            Start = start;
            End = end;
            Style = style;
        }

        internal int Start { get; }
        internal int End { get; }
        internal HtmlRenderBoxStyle Style { get; }
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
