using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutTable(IElement table, double containingWidth, HtmlRenderBoxStyle style, int depth) {
        string source = HtmlRenderStyleResolver.DescribeSource(table);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double tableWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        IReadOnlyList<IElement> rows = table.QuerySelectorAll("tr").Where(row => BelongsToTable(row, table)).ToList();
        int columnCount = rows.Count == 0 ? 0 : rows.Max(CountColumns);
        if (columnCount == 0) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.EmptyTable, "A table contained no renderable rows or cells.", HtmlDiagnosticSeverity.Info, source);
            double emptyHeight = style.MarginTop + Math.Max(1D, style.VerticalInsets) + style.MarginBottom;
            return new HtmlRenderFlowBlock(containingWidth, emptyHeight, Array.Empty<HtmlRenderVisual>(), style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, source);
        }

        double columnWidth = contentWidth / columnCount;
        var rowLayouts = new List<TableRowLayout>();
        foreach (IElement row in rows) {
            var cells = row.Children.Where(IsTableCell).ToList();
            var cellLayouts = new List<TableCellLayout>();
            int column = 0;
            double rowHeight = 0D;
            foreach (IElement cell in cells) {
                int span = ReadSpan(cell.GetAttribute("colspan"), columnCount - column);
                if (!string.IsNullOrWhiteSpace(cell.GetAttribute("rowspan")) && cell.GetAttribute("rowspan") != "1") {
                    AddUnsupported(HtmlRenderDiagnosticCodes.TableRowSpanPending, "Table row spans are not yet fragmented by the direct renderer.", cell, cell.GetAttribute("rowspan"));
                }

                double cellOuterWidth = columnWidth * span;
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, cellOuterWidth, style);
                if (cellStyle.PaddingTop == 0D && cellStyle.PaddingRight == 0D && cellStyle.PaddingBottom == 0D && cellStyle.PaddingLeft == 0D) {
                    cellStyle.PaddingTop = cellStyle.PaddingRight = cellStyle.PaddingBottom = cellStyle.PaddingLeft = 2D;
                }

                if (cellStyle.BorderWidth <= 0D) {
                    cellStyle.BorderWidth = style.BorderWidth > 0D ? style.BorderWidth : 1D;
                    cellStyle.BorderColor = style.BorderWidth > 0D ? style.BorderColor : OfficeColor.FromRgb(160, 160, 160);
                }

                double cellContentWidth = Math.Max(1D, cellOuterWidth - cellStyle.HorizontalInsets);
                HtmlInlineLayout inline = LayoutInlineNodes(cell.ChildNodes, cellContentWidth, cellStyle, depth + 1, null);
                double cellHeight = Math.Max(cellStyle.LineHeight, inline.Height) + cellStyle.VerticalInsets;
                rowHeight = Math.Max(rowHeight, cellHeight);
                cellLayouts.Add(new TableCellLayout(cell, cellStyle, inline, column, span, cellOuterWidth));
                column += span;
                if (column >= columnCount) break;
            }

            rowLayouts.Add(new TableRowLayout(cellLayouts, Math.Max(1D, rowHeight)));
        }

        double rowsHeight = rowLayouts.Sum(row => row.Height);
        double tableHeight = style.VerticalInsets + rowsHeight;
        var visuals = new List<HtmlRenderVisual>();
        var breakOffsets = new List<double>();
        AddBoxShape(visuals, style, style.MarginLeft, style.MarginTop, tableWidth, tableHeight, table);
        double contentX = style.MarginLeft + style.BorderWidth + style.PaddingLeft;
        double rowY = style.MarginTop + style.BorderWidth + style.PaddingTop;
        foreach (TableRowLayout row in rowLayouts) {
            foreach (TableCellLayout cell in row.Cells) {
                double cellX = contentX + cell.Column * columnWidth;
                OfficeShape box = OfficeShape.Rectangle(cell.Width, row.Height);
                box.FillColor = cell.Style.BackgroundColor;
                box.StrokeColor = cell.Style.BorderColor;
                box.StrokeWidth = cell.Style.BorderWidth;
                visuals.Add(new HtmlRenderShape(box, cellX, rowY, visuals.Count, source: HtmlRenderStyleResolver.DescribeSource(cell.Element)));
                double textX = cellX + cell.Style.BorderWidth + cell.Style.PaddingLeft;
                double textY = rowY + cell.Style.BorderWidth + cell.Style.PaddingTop;
                foreach (HtmlRenderVisual visual in cell.Inline.Visuals) {
                    visuals.Add(visual.Translate(textX, textY, visuals.Count));
                }
            }

            rowY += row.Height;
            breakOffsets.Add(rowY);
        }

        double outerHeight = style.MarginTop + tableHeight + style.MarginBottom;
        breakOffsets.Add(outerHeight);
        return new HtmlRenderFlowBlock(containingWidth, outerHeight, visuals, style.BreakBefore, style.BreakAfter, true, source, breakOffsets);
    }

    private static bool BelongsToTable(IElement row, IElement table) {
        IElement? current = row.ParentElement;
        while (current != null && !string.Equals(current.TagName, "table", StringComparison.OrdinalIgnoreCase)) current = current.ParentElement;
        return ReferenceEquals(current, table);
    }

    private static int CountColumns(IElement row) {
        int count = 0;
        foreach (IElement cell in row.Children.Where(IsTableCell)) count += ReadSpan(cell.GetAttribute("colspan"), int.MaxValue - count);
        return count;
    }

    private static bool IsTableCell(IElement element) => string.Equals(element.TagName, "td", StringComparison.OrdinalIgnoreCase) || string.Equals(element.TagName, "th", StringComparison.OrdinalIgnoreCase);

    private static int ReadSpan(string? value, int maximum) {
        if (!int.TryParse(value, out int span) || span <= 0) span = 1;
        return Math.Max(1, Math.Min(span, maximum));
    }

    private sealed class TableRowLayout {
        internal TableRowLayout(IReadOnlyList<TableCellLayout> cells, double height) {
            Cells = cells;
            Height = height;
        }

        internal IReadOnlyList<TableCellLayout> Cells { get; }
        internal double Height { get; }
    }

    private sealed class TableCellLayout {
        internal TableCellLayout(IElement element, HtmlRenderBoxStyle style, HtmlInlineLayout inline, int column, int span, double width) {
            Element = element;
            Style = style;
            Inline = inline;
            Column = column;
            Span = span;
            Width = width;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal HtmlInlineLayout Inline { get; }
        internal int Column { get; }
        internal int Span { get; }
        internal double Width { get; }
    }
}
