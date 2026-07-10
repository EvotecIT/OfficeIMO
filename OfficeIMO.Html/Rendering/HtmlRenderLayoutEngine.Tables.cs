using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutTable(IElement table, double containingWidth, HtmlRenderBoxStyle style, int depth) {
        string source = HtmlRenderStyleResolver.DescribeSource(table);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double tableWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        IReadOnlyList<IElement> sourceRows = table.QuerySelectorAll("tr").Where(row => BelongsToTable(row, table)).ToList();
        IReadOnlyList<IElement> rows = sourceRows.Where(row => IsHeaderRow(row, table))
            .Concat(sourceRows.Where(row => !IsHeaderRow(row, table) && !IsFooterRow(row, table)))
            .Concat(sourceRows.Where(row => IsFooterRow(row, table)))
            .ToList();
        int columnCount = DetermineColumnCount(rows, table);
        if (columnCount == 0) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.EmptyTable, "A table contained no renderable rows or cells.", HtmlDiagnosticSeverity.Info, source);
            double emptyHeight = style.MarginTop + Math.Max(1D, style.VerticalInsets) + style.MarginBottom;
            return new HtmlRenderFlowBlock(containingWidth, emptyHeight, Array.Empty<HtmlRenderVisual>(), style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, source, pageName: style.PageName);
        }

        double columnWidth = contentWidth / columnCount;
        var rowLayouts = new List<TableRowLayout>();
        var occupiedColumns = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            IElement row = rows[rowIndex];
            var cells = row.Children.Where(IsTableCell).ToList();
            var cellLayouts = new List<TableCellLayout>();
            int column = 0;
            double rowHeight = 0D;
            foreach (IElement cell in cells) {
                int requestedColumnSpan = ReadSpan(cell.GetAttribute("colspan"), 1000);
                column = FindAvailableColumn(occupiedColumns, column, requestedColumnSpan);
                if (column >= columnCount) break;
                int columnSpan = Math.Max(1, Math.Min(requestedColumnSpan, columnCount - column));
                int rowSpan = ReadRowSpan(cell.GetAttribute("rowspan"), rows, rowIndex, table);

                double cellOuterWidth = columnWidth * columnSpan;
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, cellOuterWidth, style);
                if (cellStyle.PaddingTop == 0D && cellStyle.PaddingRight == 0D && cellStyle.PaddingBottom == 0D && cellStyle.PaddingLeft == 0D) {
                    cellStyle.PaddingTop = cellStyle.PaddingRight = cellStyle.PaddingBottom = cellStyle.PaddingLeft = 2D;
                }

                if (!cellStyle.HasBorderLayout && !cellStyle.BorderDeclared) {
                    cellStyle.Borders = style.HasBorderLayout
                        ? style.Borders
                        : HtmlRenderBorderEdges.Uniform(1D, "solid", OfficeColor.FromRgb(160, 160, 160));
                }

                double cellContentWidth = Math.Max(1D, cellOuterWidth - cellStyle.HorizontalInsets);
                HtmlInlineLayout inline = LayoutInlineNodes(cell.ChildNodes, cellContentWidth, cellStyle, depth + 1, null, cell);
                double cellHeight = Math.Max(cellStyle.LineHeight, inline.Height) + cellStyle.VerticalInsets;
                if (rowSpan == 1) rowHeight = Math.Max(rowHeight, cellHeight);
                cellLayouts.Add(new TableCellLayout(cell, cellStyle, inline, column, columnSpan, rowSpan, cellOuterWidth, cellHeight));
                for (int occupiedColumn = column; occupiedColumn < column + columnSpan; occupiedColumn++) {
                    occupiedColumns[occupiedColumn] = Math.Max(occupiedColumns[occupiedColumn], rowSpan);
                }

                column += columnSpan;
                if (column >= columnCount) break;
            }

            rowLayouts.Add(new TableRowLayout(cellLayouts, Math.Max(1D, rowHeight), IsHeaderRow(row, table), IsFooterRow(row, table)));
            DecrementOccupancy(occupiedColumns);
        }

        ResolveSpanningRowHeights(rowLayouts);

        double rowsHeight = rowLayouts.Sum(row => row.Height);
        double tableHeight = style.VerticalInsets + rowsHeight;
        var visuals = new List<HtmlRenderVisual>();
        var breakOffsets = new List<double>();
        var continuationVisuals = new List<HtmlRenderVisual>();
        var trailingVisuals = new List<HtmlRenderVisual>();
        double continuationHeight = 0D;
        double trailingStart = 0D;
        double trailingHeight = 0D;
        bool collectingLeadingHeaders = true;
        IReadOnlyList<bool> canBreakAfterRows = ResolveRowBreakAvailability(rowLayouts);
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, tableWidth, tableHeight, table);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double rowY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        double headerStart = rowY;
        for (int rowIndex = 0; rowIndex < rowLayouts.Count; rowIndex++) {
            TableRowLayout row = rowLayouts[rowIndex];
            int rowVisualStart = visuals.Count;
            if (row.IsFooter && trailingVisuals.Count == 0) trailingStart = rowY;
            foreach (TableCellLayout cell in row.Cells) {
                double cellX = contentX + cell.Column * columnWidth;
                double cellHeight = GetSpanningHeight(rowLayouts, rowIndex, cell.RowSpan);
                AddBoxPaint(visuals, cell.Style, cellX, rowY, cell.Width, cellHeight, cell.Element);
                double textX = cellX + cell.Style.BorderLeftWidth + cell.Style.PaddingLeft;
                double textY = rowY + cell.Style.BorderTopWidth + cell.Style.PaddingTop;
                foreach (HtmlRenderVisual visual in cell.Inline.Visuals) {
                    visuals.Add(visual.Translate(textX, textY, visuals.Count));
                }
                AddBoxOutlinePaint(visuals, cell.Style, cellX, rowY, cell.Width, cellHeight, cell.Element);
            }

            if (collectingLeadingHeaders && row.IsHeader) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    continuationVisuals.Add(visuals[visualIndex].Translate(0D, -headerStart, continuationVisuals.Count));
                }

                continuationHeight += row.Height;
            } else {
                collectingLeadingHeaders = false;
            }

            if (row.IsFooter) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    trailingVisuals.Add(visuals[visualIndex].Translate(0D, -trailingStart, trailingVisuals.Count));
                }

                trailingHeight += row.Height;
            }

            rowY += row.Height;
            bool headerHasBodyAfter = row.IsHeader && rowLayouts.Skip(rowIndex + 1).Any(candidate => !candidate.IsHeader && !candidate.IsFooter);
            if (!headerHasBodyAfter && canBreakAfterRows[rowIndex]) breakOffsets.Add(rowY);
        }
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, tableWidth, tableHeight, table);

        double outerHeight = style.MarginTop + tableHeight + style.MarginBottom;
        breakOffsets.Add(outerHeight);
        IEnumerable<HtmlRenderTrailingGroup> trailingGroups = trailingVisuals.Count > 0 && trailingHeight > 0D
            ? new[] { new HtmlRenderTrailingGroup(0D, trailingStart, outerHeight, outerHeight - trailingStart, trailingVisuals) }
            : Array.Empty<HtmlRenderTrailingGroup>();
        return new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            true,
            source,
            breakOffsets,
            trailingGroups: trailingGroups,
            continuationVisuals: continuationVisuals,
            continuationHeight: continuationHeight,
            continuationStartsAfter: headerStart + continuationHeight,
            pageName: style.PageName);
    }

    private static bool BelongsToTable(IElement row, IElement table) {
        IElement? current = row.ParentElement;
        while (current != null && !string.Equals(current.TagName, "table", StringComparison.OrdinalIgnoreCase)) current = current.ParentElement;
        return ReferenceEquals(current, table);
    }

    private static int DetermineColumnCount(IReadOnlyList<IElement> rows, IElement table) {
        var occupancy = new List<int>();
        int maximum = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (IElement cell in rows[rowIndex].Children.Where(IsTableCell)) {
                int columnSpan = ReadSpan(cell.GetAttribute("colspan"), 1000);
                column = FindAvailableColumn(occupancy, column, columnSpan);
                EnsureOccupancySize(occupancy, column + columnSpan);
                int rowSpan = ReadRowSpan(cell.GetAttribute("rowspan"), rows, rowIndex, table);
                for (int occupiedColumn = column; occupiedColumn < column + columnSpan; occupiedColumn++) {
                    occupancy[occupiedColumn] = Math.Max(occupancy[occupiedColumn], rowSpan);
                }

                column += columnSpan;
                maximum = Math.Max(maximum, column);
            }

            DecrementOccupancy(occupancy);
        }

        return maximum;
    }

    private static int FindAvailableColumn(IReadOnlyList<int> occupancy, int start, int span) {
        int column = Math.Max(0, start);
        while (column < occupancy.Count) {
            bool available = true;
            for (int offset = 0; offset < span && column + offset < occupancy.Count; offset++) {
                if (occupancy[column + offset] <= 0) continue;
                column += offset + 1;
                available = false;
                break;
            }

            if (available) return column;
        }

        return column;
    }

    private static void EnsureOccupancySize(List<int> occupancy, int size) {
        while (occupancy.Count < size) occupancy.Add(0);
    }

    private static void DecrementOccupancy(IList<int> occupancy) {
        for (int column = 0; column < occupancy.Count; column++) {
            if (occupancy[column] > 0) occupancy[column]--;
        }
    }

    private static int ReadRowSpan(string? value, IReadOnlyList<IElement> rows, int rowIndex, IElement table) {
        int maximum = CountRowsRemainingInGroup(rows, rowIndex, table);
        if (!int.TryParse(value, out int requested) || requested < 0) return 1;
        if (requested == 0) return maximum;
        return Math.Max(1, Math.Min(requested, maximum));
    }

    private static int CountRowsRemainingInGroup(IReadOnlyList<IElement> rows, int rowIndex, IElement table) {
        IElement group = GetRowGroup(rows[rowIndex], table);
        int count = 0;
        for (int index = rowIndex; index < rows.Count && ReferenceEquals(GetRowGroup(rows[index], table), group); index++) count++;
        return Math.Max(1, count);
    }

    private static IElement GetRowGroup(IElement row, IElement table) {
        IElement? parent = row.ParentElement;
        return parent == null || ReferenceEquals(parent, table) ? table : parent;
    }

    private static void ResolveSpanningRowHeights(IReadOnlyList<TableRowLayout> rows) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells.Where(cell => cell.RowSpan > 1)) {
                double currentHeight = GetSpanningHeight(rows, rowIndex, cell.RowSpan);
                double deficit = cell.MinimumHeight - currentHeight;
                if (deficit <= 0.0001D) continue;
                double addition = deficit / cell.RowSpan;
                for (int offset = 0; offset < cell.RowSpan && rowIndex + offset < rows.Count; offset++) rows[rowIndex + offset].Height += addition;
            }
        }
    }

    private static double GetSpanningHeight(IReadOnlyList<TableRowLayout> rows, int rowIndex, int rowSpan) {
        double height = 0D;
        for (int offset = 0; offset < rowSpan && rowIndex + offset < rows.Count; offset++) height += rows[rowIndex + offset].Height;
        return height;
    }

    private static IReadOnlyList<bool> ResolveRowBreakAvailability(IReadOnlyList<TableRowLayout> rows) {
        var result = new bool[rows.Count];
        int occupiedThrough = -1;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells) {
                occupiedThrough = Math.Max(occupiedThrough, rowIndex + cell.RowSpan - 1);
            }

            result[rowIndex] = occupiedThrough <= rowIndex;
        }

        return result;
    }

    private static bool IsTableCell(IElement element) => string.Equals(element.TagName, "td", StringComparison.OrdinalIgnoreCase) || string.Equals(element.TagName, "th", StringComparison.OrdinalIgnoreCase);

    private static bool IsHeaderRow(IElement row, IElement table) {
        IElement? current = row.ParentElement;
        while (current != null && !ReferenceEquals(current, table)) {
            if (string.Equals(current.TagName, "thead", StringComparison.OrdinalIgnoreCase)) return true;
            current = current.ParentElement;
        }

        return false;
    }

    private static bool IsFooterRow(IElement row, IElement table) {
        IElement? current = row.ParentElement;
        while (current != null && !ReferenceEquals(current, table)) {
            if (string.Equals(current.TagName, "tfoot", StringComparison.OrdinalIgnoreCase)) return true;
            current = current.ParentElement;
        }

        return false;
    }

    private static int ReadSpan(string? value, int maximum) {
        if (!int.TryParse(value, out int span) || span <= 0) span = 1;
        return Math.Max(1, Math.Min(span, maximum));
    }

    private sealed class TableRowLayout {
        internal TableRowLayout(IReadOnlyList<TableCellLayout> cells, double height, bool isHeader, bool isFooter) {
            Cells = cells;
            Height = height;
            IsHeader = isHeader;
            IsFooter = isFooter;
        }

        internal IReadOnlyList<TableCellLayout> Cells { get; }
        internal double Height { get; set; }
        internal bool IsHeader { get; }
        internal bool IsFooter { get; }
    }

    private sealed class TableCellLayout {
        internal TableCellLayout(IElement element, HtmlRenderBoxStyle style, HtmlInlineLayout inline, int column, int span, int rowSpan, double width, double minimumHeight) {
            Element = element;
            Style = style;
            Inline = inline;
            Column = column;
            Span = span;
            RowSpan = rowSpan;
            Width = width;
            MinimumHeight = minimumHeight;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal HtmlInlineLayout Inline { get; }
        internal int Column { get; }
        internal int Span { get; }
        internal int RowSpan { get; }
        internal double Width { get; }
        internal double MinimumHeight { get; }
    }
}
