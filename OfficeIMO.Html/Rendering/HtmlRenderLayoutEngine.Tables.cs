using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderFlowBlock LayoutTable(IElement table, double containingWidth, HtmlRenderBoxStyle style, int depth) {
        string source = HtmlRenderStyleResolver.DescribeSource(table);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double tableWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, tableWidth - style.HorizontalInsets);
        TableCaptionLayout? caption = LayoutTableCaption(table, tableWidth, style, depth);
        double topCaptionHeight = caption != null && caption.Side == "top" ? caption.Height : 0D;
        double bottomCaptionHeight = caption != null && caption.Side == "bottom" ? caption.Height : 0D;
        double tableY = style.MarginTop + topCaptionHeight;
        ReportUnsupportedTableValues(table, style);
        IReadOnlyList<IElement> sourceRows = table.QuerySelectorAll("tr").Where(row => BelongsToTable(row, table)).ToList();
        if (sourceRows.Count > _options.MaxTableRows) {
            throw new HtmlDomLimitException(
                HtmlRenderDiagnosticCodes.TableLimitExceeded,
                "HTML table row count exceeded the configured maximum.",
                nameof(HtmlRenderOptions.MaxTableRows),
                sourceRows.Count,
                _options.MaxTableRows);
        }
        var rowGroupStyles = new Dictionary<IElement, HtmlRenderBoxStyle>();
        var rowStyles = new Dictionary<IElement, HtmlRenderBoxStyle>();
        var renderableRows = new List<IElement>();
        foreach (IElement row in sourceRows) {
            IElement rowGroup = GetRowGroup(row, table);
            HtmlRenderBoxStyle rowParentStyle = style;
            if (!ReferenceEquals(rowGroup, table)) {
                if (!rowGroupStyles.TryGetValue(rowGroup, out HtmlRenderBoxStyle? rowGroupStyle)) {
                    rowGroupStyle = _styleResolver.Resolve(rowGroup, contentWidth, style);
                    rowGroupStyles[rowGroup] = rowGroupStyle;
                    _layoutStyles[rowGroup] = rowGroupStyle.Clone();
                }

                if (rowGroupStyle.Display == "none") continue;
                rowParentStyle = rowGroupStyle;
            }

            HtmlRenderBoxStyle rowStyle = _styleResolver.Resolve(row, contentWidth, rowParentStyle);
            rowStyles[row] = rowStyle;
            _layoutStyles[row] = rowStyle.Clone();
            if (rowStyle.Display != "none") renderableRows.Add(row);
        }

        IReadOnlyList<IElement> rows = renderableRows.Where(row => IsHeaderRow(row, table))
            .Concat(renderableRows.Where(row => !IsHeaderRow(row, table) && !IsFooterRow(row, table)))
            .Concat(renderableRows.Where(row => IsFooterRow(row, table)))
            .ToList();
        int rowColumnCount = DetermineColumnCount(rows, table);
        if (rowColumnCount == 0) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.EmptyTable, "A table contained no renderable rows or cells.", HtmlDiagnosticSeverity.Info, source);
            double emptyTableHeight = Math.Max(1D, style.VerticalInsets);
            var emptyVisuals = new List<HtmlRenderVisual>();
            if (caption != null && caption.Side == "top") AppendTableCaption(emptyVisuals, caption, style.MarginLeft, style.MarginTop);
            AddBoxPaint(emptyVisuals, style, style.MarginLeft, tableY, tableWidth, emptyTableHeight, table);
            AddBoxOutlinePaint(emptyVisuals, style, style.MarginLeft, tableY, tableWidth, emptyTableHeight, table);
            if (caption != null && caption.Side == "bottom") AppendTableCaption(emptyVisuals, caption, style.MarginLeft, tableY + emptyTableHeight);
            double emptyHeight = style.MarginTop + topCaptionHeight + emptyTableHeight + bottomCaptionHeight + style.MarginBottom;
            IReadOnlyList<HtmlRenderVisual> semanticEmptyVisuals = new[] {
                new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Table,
                    style.MarginLeft,
                    style.MarginTop,
                    tableWidth,
                    Math.Max(0.01D, topCaptionHeight + emptyTableHeight + bottomCaptionHeight),
                    emptyVisuals,
                    0,
                    source)
            };
            return new HtmlRenderFlowBlock(containingWidth, emptyHeight, semanticEmptyVisuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, source, pageName: style.PageName);
        }

        int columnCount = Math.Max(rowColumnCount, DetermineDeclaredColumnCount(table));
        double horizontalSpacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingX;
        double verticalSpacing = style.BorderCollapse == "collapse" ? 0D : style.BorderSpacingY;
        double trackWidth = Math.Max(0.01D, contentWidth - horizontalSpacing * (columnCount + 1));
        IReadOnlyList<double> columnWidths = ResolveTableColumnWidths(rows, table, columnCount, trackWidth, style);
        double[] columnOffsets = CreateColumnOffsets(columnWidths);
        var rowLayouts = new List<TableRowLayout>();
        var occupiedColumns = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            CheckCancellation();
            IElement row = rows[rowIndex];
            IElement rowGroup = GetRowGroup(row, table);
            IElement? rowGroupElement = null;
            HtmlRenderBoxStyle? rowGroupStyle = null;
            if (!ReferenceEquals(rowGroup, table)) {
                rowGroupElement = rowGroup;
                rowGroupStyle = rowGroupStyles[rowGroup];
            }
            HtmlRenderBoxStyle rowStyle = rowStyles[row];
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

                double cellOuterWidth = SumColumnWidths(columnWidths, column, columnSpan) + horizontalSpacing * (columnSpan - 1);
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, cellOuterWidth, style);
                if (cellStyle.PaddingTop == 0D && cellStyle.PaddingRight == 0D && cellStyle.PaddingBottom == 0D && cellStyle.PaddingLeft == 0D) {
                    cellStyle.PaddingTop = cellStyle.PaddingRight = cellStyle.PaddingBottom = cellStyle.PaddingLeft = 2D;
                }

                if (!cellStyle.HasBorderLayout && !cellStyle.BorderDeclared) {
                    cellStyle.Borders = style.BorderCollapse == "collapse" && style.HasBorderLayout
                        ? HtmlRenderBorderEdges.Uniform(0D, "none", cellStyle.Color)
                        : style.HasBorderLayout
                            ? style.Borders
                            : HtmlRenderBorderEdges.Uniform(1D, "solid", OfficeColor.FromRgb(160, 160, 160));
                }

                double cellContentWidth = Math.Max(1D, cellOuterWidth - cellStyle.HorizontalInsets);
                HtmlInlineLayout inline = LayoutTableCellContent(cell, cellContentWidth, cellStyle, depth + 1);
                double cellHeight = Math.Max(cellStyle.LineHeight, inline.Height) + cellStyle.VerticalInsets;
                if (rowSpan == 1) rowHeight = Math.Max(rowHeight, cellHeight);
                cellLayouts.Add(new TableCellLayout(cell, cellStyle, inline, column, columnSpan, rowSpan, cellOuterWidth, cellHeight));
                for (int occupiedColumn = column; occupiedColumn < column + columnSpan; occupiedColumn++) {
                    occupiedColumns[occupiedColumn] = Math.Max(occupiedColumns[occupiedColumn], rowSpan);
                }

                column += columnSpan;
                if (column >= columnCount) break;
            }

            rowLayouts.Add(new TableRowLayout(
                row,
                rowStyle,
                rowGroupElement,
                rowGroupStyle,
                cellLayouts,
                Math.Max(1D, rowHeight),
                IsHeaderRow(row, table),
                IsFooterRow(row, table)));
            DecrementOccupancy(occupiedColumns);
        }

        ResolveSpanningRowHeights(rowLayouts, verticalSpacing);

        double rowsHeight = rowLayouts.Sum(row => row.Height) + verticalSpacing * (rowLayouts.Count + 1);
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
        if (caption != null && caption.Side == "top") AppendTableCaption(visuals, caption, style.MarginLeft, style.MarginTop);
        HtmlRenderBoxStyle tablePaintStyle = style.BorderCollapse == "collapse" ? CreateCollapsedCellPaintStyle(style) : style;
        AddBoxPaint(visuals, tablePaintStyle, style.MarginLeft, tableY, tableWidth, tableHeight, table);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double rowY = tableY + style.BorderTopWidth + style.PaddingTop + verticalSpacing;
        double headerStart = rowY;
        for (int rowIndex = 0; rowIndex < rowLayouts.Count; rowIndex++) {
            CheckCancellation();
            TableRowLayout row = rowLayouts[rowIndex];
            int rowVisualStart = visuals.Count;
            if (row.IsFooter && trailingVisuals.Count == 0) trailingStart = rowY;
            var rowVisuals = new List<HtmlRenderVisual>();
            double rowPaintX = contentX + horizontalSpacing;
            double rowPaintWidth = Math.Max(0.01D, contentWidth - horizontalSpacing * 2D);
            bool startsRowGroup = row.GroupElement != null
                && (rowIndex == 0 || !ReferenceEquals(rowLayouts[rowIndex - 1].GroupElement, row.GroupElement));
            if (startsRowGroup && row.GroupStyle != null) {
                int groupRowCount = rowLayouts.Skip(rowIndex).TakeWhile(candidate => ReferenceEquals(candidate.GroupElement, row.GroupElement)).Count();
                double groupHeight = rowLayouts.Skip(rowIndex).Take(groupRowCount).Sum(candidate => candidate.Height)
                    + verticalSpacing * Math.Max(0, groupRowCount - 1);
                string groupSource = HtmlRenderStyleResolver.DescribeSource(row.GroupElement!);
                AddBoxBackground(rowVisuals, row.GroupStyle, rowPaintX, rowY, rowPaintWidth, groupHeight, 0D, row.GroupElement!, groupSource, groupSource);
            }
            string rowSource = HtmlRenderStyleResolver.DescribeSource(row.Element);
            AddBoxBackground(rowVisuals, row.Style, rowPaintX, rowY, rowPaintWidth, row.Height, 0D, row.Element, rowSource, rowSource);
            foreach (TableCellLayout cell in row.Cells) {
                double cellX = contentX + horizontalSpacing + columnOffsets[cell.Column] + horizontalSpacing * cell.Column;
                double cellHeight = GetSpanningHeight(rowLayouts, rowIndex, cell.RowSpan, verticalSpacing);
                HtmlRenderBoxStyle paintStyle = style.BorderCollapse == "collapse" ? CreateCollapsedCellPaintStyle(cell.Style) : cell.Style;
                var cellVisuals = new List<HtmlRenderVisual>();
                AddBoxPaint(cellVisuals, paintStyle, cellX, rowY, cell.Width, cellHeight, cell.Element);
                double textX = cellX + cell.Style.BorderLeftWidth + cell.Style.PaddingLeft;
                double textY = rowY + cell.Style.BorderTopWidth + cell.Style.PaddingTop;
                foreach (HtmlRenderVisual visual in cell.Inline.Visuals) {
                    cellVisuals.Add(visual.Translate(textX, textY, cellVisuals.Count));
                }
                AddBoxOutlinePaint(cellVisuals, cell.Style, cellX, rowY, cell.Width, cellHeight, cell.Element);
                bool headerCell = string.Equals(cell.Element.TagName, "th", StringComparison.OrdinalIgnoreCase);
                rowVisuals.Add(new HtmlRenderSemanticGroup(
                    headerCell ? HtmlRenderSemanticGroupRole.TableHeaderCell : HtmlRenderSemanticGroupRole.TableCell,
                    cellX,
                    rowY,
                    cell.Width,
                    cellHeight,
                    cellVisuals,
                    rowVisuals.Count,
                    HtmlRenderStyleResolver.DescribeSource(cell.Element),
                    cell.Span,
                    cell.RowSpan,
                    headerCell ? ResolveTableHeaderScope(cell.Element) : null));
            }

            visuals.Add(new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.TableRow,
                contentX,
                rowY,
                contentWidth,
                row.Height,
                rowVisuals,
                visuals.Count,
                HtmlRenderStyleResolver.DescribeSource(row.Element)));

            if (!row.IsHeader && !row.IsFooter && row.Cells.Count == 1 && row.Cells[0].RowSpan == 1) {
                TableCellLayout cell = row.Cells[0];
                double textY = rowY + cell.Style.BorderTopWidth + cell.Style.PaddingTop;
                breakOffsets.AddRange(cell.Inline.BreakOffsets
                    .Where(offset => offset <= cell.Inline.Height - cell.Style.LineHeight + 0.0001D)
                    .Select(offset => textY + offset));
            }

            if (collectingLeadingHeaders && row.IsHeader) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    continuationVisuals.Add(visuals[visualIndex].Translate(0D, -headerStart, continuationVisuals.Count));
                }

                continuationHeight += row.Height + verticalSpacing;
            } else {
                collectingLeadingHeaders = false;
            }

            if (row.IsFooter) {
                for (int visualIndex = rowVisualStart; visualIndex < visuals.Count; visualIndex++) {
                    trailingVisuals.Add(visuals[visualIndex].Translate(0D, -trailingStart, trailingVisuals.Count));
                }

                trailingHeight += row.Height + verticalSpacing;
            }

            rowY += row.Height + verticalSpacing;
            bool headerHasBodyAfter = row.IsHeader && rowLayouts.Skip(rowIndex + 1).Any(candidate => !candidate.IsHeader && !candidate.IsFooter);
            if (!headerHasBodyAfter && canBreakAfterRows[rowIndex]) breakOffsets.Add(rowY);
        }
        if (style.BorderCollapse == "collapse") {
            AddCollapsedTableBorders(visuals, table, style, rowLayouts, columnWidths, columnOffsets, contentX, tableY + style.BorderTopWidth + style.PaddingTop);
        }
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, tableY, tableWidth, tableHeight, table);
        if (caption != null && caption.Side == "bottom") AppendTableCaption(visuals, caption, style.MarginLeft, tableY + tableHeight);

        double outerHeight = style.MarginTop + topCaptionHeight + tableHeight + bottomCaptionHeight + style.MarginBottom;
        breakOffsets.Add(outerHeight);
        IReadOnlyList<HtmlRenderVisual> semanticVisuals = new[] {
            new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Table,
                style.MarginLeft,
                style.MarginTop,
                tableWidth,
                Math.Max(0.01D, topCaptionHeight + tableHeight + bottomCaptionHeight),
                visuals,
                0,
                source)
        };
        IReadOnlyList<HtmlRenderVisual> semanticContinuationVisuals = continuationVisuals.Count == 0
            ? Array.Empty<HtmlRenderVisual>()
            : new[] {
                new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Table,
                    style.MarginLeft,
                    0D,
                    tableWidth,
                    Math.Max(0.01D, continuationHeight),
                    continuationVisuals,
                    0,
                    source)
            };
        IReadOnlyList<HtmlRenderVisual> semanticTrailingVisuals = trailingVisuals.Count == 0
            ? Array.Empty<HtmlRenderVisual>()
            : new[] {
                new HtmlRenderSemanticGroup(
                    HtmlRenderSemanticGroupRole.Table,
                    style.MarginLeft,
                    0D,
                    tableWidth,
                    Math.Max(0.01D, trailingHeight),
                    trailingVisuals,
                    0,
                    source)
            };
        IEnumerable<HtmlRenderTrailingGroup> trailingGroups = trailingVisuals.Count > 0 && trailingHeight > 0D
            ? new[] { new HtmlRenderTrailingGroup(0D, trailingStart, outerHeight, outerHeight - trailingStart, semanticTrailingVisuals) }
            : Array.Empty<HtmlRenderTrailingGroup>();
        return new HtmlRenderFlowBlock(
            containingWidth,
            outerHeight,
            semanticVisuals,
            style.BreakBefore,
            style.BreakAfter,
            true,
            source,
            breakOffsets,
            trailingGroups: trailingGroups,
            continuationVisuals: semanticContinuationVisuals,
            continuationHeight: continuationHeight,
            continuationStartsAfter: headerStart + continuationHeight,
            pageName: style.PageName);
    }

    private static bool BelongsToTable(IElement row, IElement table) {
        IElement? current = row.ParentElement;
        while (current != null && !string.Equals(current.TagName, "table", StringComparison.OrdinalIgnoreCase)) current = current.ParentElement;
        return ReferenceEquals(current, table);
    }

    private int DetermineColumnCount(IReadOnlyList<IElement> rows, IElement table) {
        var occupancy = new List<int>();
        int maximum = 0;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (IElement cell in rows[rowIndex].Children.Where(IsTableCell)) {
                int columnSpan = ReadSpan(cell.GetAttribute("colspan"), 1000);
                column = FindAvailableColumn(occupancy, column, columnSpan);
                long columnEnd = (long)column + columnSpan;
                EnsureTableColumnLimit(columnEnd);
                EnsureOccupancySize(occupancy, (int)columnEnd);
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

    private void EnsureTableColumnLimit(long count) {
        if (count <= _options.MaxTableColumns) return;
        throw new HtmlDomLimitException(
            HtmlRenderDiagnosticCodes.TableLimitExceeded,
            "HTML table column count exceeded the configured maximum.",
            nameof(HtmlRenderOptions.MaxTableColumns),
            count,
            _options.MaxTableColumns);
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

    private static void ResolveSpanningRowHeights(IReadOnlyList<TableRowLayout> rows, double verticalSpacing) {
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            foreach (TableCellLayout cell in rows[rowIndex].Cells.Where(cell => cell.RowSpan > 1)) {
                double currentHeight = GetSpanningHeight(rows, rowIndex, cell.RowSpan, verticalSpacing);
                double deficit = cell.MinimumHeight - currentHeight;
                if (deficit <= 0.0001D) continue;
                double addition = deficit / cell.RowSpan;
                for (int offset = 0; offset < cell.RowSpan && rowIndex + offset < rows.Count; offset++) rows[rowIndex + offset].Height += addition;
            }
        }
    }

    private static double GetSpanningHeight(IReadOnlyList<TableRowLayout> rows, int rowIndex, int rowSpan, double verticalSpacing) {
        double height = 0D;
        int includedRows = 0;
        for (int offset = 0; offset < rowSpan && rowIndex + offset < rows.Count; offset++) {
            height += rows[rowIndex + offset].Height;
            includedRows++;
        }
        if (includedRows > 1) height += verticalSpacing * (includedRows - 1);
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

    private static HtmlRenderTableHeaderScope ResolveTableHeaderScope(IElement cell) {
        string scope = cell.GetAttribute("scope")?.Trim().ToLowerInvariant() ?? string.Empty;
        if (scope == "row" || scope == "rowgroup") return HtmlRenderTableHeaderScope.Row;
        if (scope == "both") return HtmlRenderTableHeaderScope.Both;
        return HtmlRenderTableHeaderScope.Column;
    }

    private TableCaptionLayout? LayoutTableCaption(
        IElement table,
        double tableWidth,
        HtmlRenderBoxStyle tableStyle,
        int depth) {
        IElement? element = table.Children.FirstOrDefault(child => string.Equals(child.TagName, "caption", StringComparison.OrdinalIgnoreCase));
        if (element == null) return null;

        HtmlRenderBoxStyle style = _styleResolver.Resolve(element, tableWidth, tableStyle);
        _layoutStyles[element] = style.Clone();
        if (tableStyle.UnsupportedCaptionSide.Length == 0) ReportUnsupportedTableValues(element, style);
        if (style.Display == "none") return null;

        double availableWidth = Math.Max(1D, tableWidth - style.MarginLeft - style.MarginRight);
        double boxWidth = ResolveBoxWidth(availableWidth, style);
        double contentWidth = Math.Max(1D, boxWidth - style.HorizontalInsets);
        HtmlInlineLayout inline = LayoutInlineNodes(element.ChildNodes, contentWidth, style, depth + 1, null, element);
        double contentHeight = Math.Max(style.LineHeight, inline.Height);
        double boxHeight = ResolveBoxHeight(contentHeight, style);
        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
        foreach (HtmlRenderVisual visual in inline.Visuals) visuals.Add(visual.Translate(contentX, contentY, visuals.Count));
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, style.MarginTop, boxWidth, boxHeight, element);
        IReadOnlyList<HtmlRenderVisual> semanticVisuals = new[] {
            new HtmlRenderSemanticGroup(
                HtmlRenderSemanticGroupRole.Caption,
                style.MarginLeft,
                style.MarginTop,
                boxWidth,
                boxHeight,
                visuals,
                0,
                HtmlRenderStyleResolver.DescribeSource(element))
        };
        return new TableCaptionLayout(style.CaptionSide, style.MarginTop + boxHeight + style.MarginBottom, semanticVisuals);
    }

    private void ReportUnsupportedTableValues(IElement element, HtmlRenderBoxStyle style) {
        var details = new List<string>(2);
        if (style.UnsupportedCaptionSide.Length > 0) details.Add("caption-side=" + style.UnsupportedCaptionSide);
        if (style.UnsupportedTableLayout.Length > 0) details.Add("table-layout=" + style.UnsupportedTableLayout);
        if (style.UnsupportedBorderCollapse.Length > 0) details.Add("border-collapse=" + style.UnsupportedBorderCollapse);
        if (style.UnsupportedBorderSpacing.Length > 0) details.Add("border-spacing=" + style.UnsupportedBorderSpacing);
        if (details.Count == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.TableValueUnsupported,
            "An unsupported table formatting value used its documented fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }

    private static void AppendTableCaption(
        ICollection<HtmlRenderVisual> target,
        TableCaptionLayout caption,
        double x,
        double y) {
        foreach (HtmlRenderVisual visual in caption.Visuals) target.Add(visual.Translate(x, y, target.Count));
    }

    private sealed class TableCaptionLayout {
        internal TableCaptionLayout(string side, double height, IReadOnlyList<HtmlRenderVisual> visuals) {
            Side = side;
            Height = height;
            Visuals = visuals;
        }

        internal string Side { get; }
        internal double Height { get; }
        internal IReadOnlyList<HtmlRenderVisual> Visuals { get; }
    }

    private sealed class TableRowLayout {
        internal TableRowLayout(
            IElement element,
            HtmlRenderBoxStyle style,
            IElement? groupElement,
            HtmlRenderBoxStyle? groupStyle,
            IReadOnlyList<TableCellLayout> cells,
            double height,
            bool isHeader,
            bool isFooter) {
            Element = element;
            Style = style;
            GroupElement = groupElement;
            GroupStyle = groupStyle;
            Cells = cells;
            Height = height;
            IsHeader = isHeader;
            IsFooter = isFooter;
        }

        internal IElement Element { get; }
        internal HtmlRenderBoxStyle Style { get; }
        internal IElement? GroupElement { get; }
        internal HtmlRenderBoxStyle? GroupStyle { get; }
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
