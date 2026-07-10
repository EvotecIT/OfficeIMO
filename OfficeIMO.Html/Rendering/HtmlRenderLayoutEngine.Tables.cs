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
        IReadOnlyList<IElement> rows = sourceRows.Where(row => IsHeaderRow(row, table))
            .Concat(sourceRows.Where(row => !IsHeaderRow(row, table) && !IsFooterRow(row, table)))
            .Concat(sourceRows.Where(row => IsFooterRow(row, table)))
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
            return new HtmlRenderFlowBlock(containingWidth, emptyHeight, emptyVisuals, style.BreakBefore, style.BreakAfter, style.AvoidBreakInside, source, pageName: style.PageName);
        }

        int columnCount = Math.Max(rowColumnCount, DetermineDeclaredColumnCount(table));
        IReadOnlyList<double> columnWidths = ResolveTableColumnWidths(rows, table, columnCount, contentWidth, style);
        double[] columnOffsets = CreateColumnOffsets(columnWidths);
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

                double cellOuterWidth = SumColumnWidths(columnWidths, column, columnSpan);
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
        if (caption != null && caption.Side == "top") AppendTableCaption(visuals, caption, style.MarginLeft, style.MarginTop);
        AddBoxPaint(visuals, style, style.MarginLeft, tableY, tableWidth, tableHeight, table);
        double contentX = style.MarginLeft + style.BorderLeftWidth + style.PaddingLeft;
        double rowY = tableY + style.BorderTopWidth + style.PaddingTop;
        double headerStart = rowY;
        for (int rowIndex = 0; rowIndex < rowLayouts.Count; rowIndex++) {
            TableRowLayout row = rowLayouts[rowIndex];
            int rowVisualStart = visuals.Count;
            if (row.IsFooter && trailingVisuals.Count == 0) trailingStart = rowY;
            foreach (TableCellLayout cell in row.Cells) {
                double cellX = contentX + columnOffsets[cell.Column];
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
        AddBoxOutlinePaint(visuals, style, style.MarginLeft, tableY, tableWidth, tableHeight, table);
        if (caption != null && caption.Side == "bottom") AppendTableCaption(visuals, caption, style.MarginLeft, tableY + tableHeight);

        double outerHeight = style.MarginTop + topCaptionHeight + tableHeight + bottomCaptionHeight + style.MarginBottom;
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
        return new TableCaptionLayout(style.CaptionSide, style.MarginTop + boxHeight + style.MarginBottom, visuals);
    }

    private void ReportUnsupportedTableValues(IElement element, HtmlRenderBoxStyle style) {
        var details = new List<string>(2);
        if (style.UnsupportedCaptionSide.Length > 0) details.Add("caption-side=" + style.UnsupportedCaptionSide);
        if (style.UnsupportedTableLayout.Length > 0) details.Add("table-layout=" + style.UnsupportedTableLayout);
        if (details.Count == 0) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.TableValueUnsupported,
            "An unsupported table formatting value used its documented fallback.",
            HtmlDiagnosticSeverity.Warning,
            HtmlRenderStyleResolver.DescribeSource(element),
            string.Join(";", details));
    }

    private IReadOnlyList<double> ResolveTableColumnWidths(
        IReadOnlyList<IElement> rows,
        IElement table,
        int columnCount,
        double contentWidth,
        HtmlRenderBoxStyle tableStyle) {
        if (tableStyle.TableLayout == "fixed") {
            var fixedWidths = new double[columnCount];
            ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, fixedWidths, fixedWidths);
            ApplyFirstRowAuthoredWidths(rows, tableStyle, fixedWidths, contentWidth);
            return AllocateFixedColumnWidths(fixedWidths, contentWidth);
        }

        var minimums = Enumerable.Repeat(1D, columnCount).ToArray();
        var preferred = Enumerable.Repeat(1D, columnCount).ToArray();
        ApplyDeclaredColumnWidths(table, contentWidth, tableStyle, minimums, preferred);

        var occupancy = new int[columnCount];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            int column = 0;
            foreach (IElement cell in rows[rowIndex].Children.Where(IsTableCell)) {
                int requestedSpan = ReadSpan(cell.GetAttribute("colspan"), columnCount);
                column = FindAvailableColumn(occupancy, column, requestedSpan);
                if (column >= columnCount) break;
                int span = Math.Max(1, Math.Min(requestedSpan, columnCount - column));
                HtmlRenderBoxStyle cellStyle = _styleResolver.Resolve(cell, contentWidth, tableStyle);
                ResolveTableCellIntrinsicWidths(cell, cellStyle, out double minimum, out double maximum);
                ApplySpanningWidth(minimums, column, span, minimum);
                ApplySpanningWidth(preferred, column, span, maximum);
                int rowSpan = ReadRowSpan(cell.GetAttribute("rowspan"), rows, rowIndex, table);
                for (int occupied = column; occupied < column + span; occupied++) occupancy[occupied] = Math.Max(occupancy[occupied], rowSpan);
                column += span;
            }
            DecrementOccupancy(occupancy);
        }

        return AllocateAutoColumnWidths(minimums, preferred, contentWidth);
    }

    private void ApplyFirstRowAuthoredWidths(
        IReadOnlyList<IElement> rows,
        HtmlRenderBoxStyle tableStyle,
        double[] widths,
        double contentWidth) {
        if (rows.Count == 0) return;
        int column = 0;
        foreach (IElement cell in rows[0].Children.Where(IsTableCell)) {
            int span = Math.Min(ReadSpan(cell.GetAttribute("colspan"), widths.Length), widths.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(cell, contentWidth, tableStyle);
            if (style.ExplicitWidth.HasValue) {
                double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
                ApplySpanningWidth(widths, column, span, authored);
            }
            column += span;
            if (column >= widths.Length) break;
        }
    }

    private void ApplyDeclaredColumnWidths(
        IElement table,
        double contentWidth,
        HtmlRenderBoxStyle tableStyle,
        double[] minimums,
        double[] preferred) {
        int column = 0;
        foreach (IElement element in table.QuerySelectorAll("col").Where(candidate => BelongsToTableColumn(candidate, table))) {
            int span = Math.Min(ReadSpan(element.GetAttribute("span"), minimums.Length), minimums.Length - column);
            if (span <= 0) break;
            HtmlRenderBoxStyle style = _styleResolver.Resolve(element, contentWidth, tableStyle);
            if (style.ExplicitWidth.HasValue) {
                double width = Math.Max(1D, style.ExplicitWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets));
                double perColumn = width / span;
                for (int offset = 0; offset < span; offset++) {
                    minimums[column + offset] = Math.Max(minimums[column + offset], perColumn);
                    preferred[column + offset] = Math.Max(preferred[column + offset], perColumn);
                }
            }
            column += span;
            if (column >= minimums.Length) break;
        }
    }

    private void ResolveTableCellIntrinsicWidths(
        IElement cell,
        HtmlRenderBoxStyle style,
        out double minimum,
        out double preferred) {
        string text = ApplyTextTransform(cell.TextContent ?? string.Empty, style.TextTransform);
        IReadOnlyList<string> tokens = HtmlRenderCssValues.SplitWhitespace(text);
        string normalized = string.Join(" ", tokens);
        double insets = style.HorizontalInsets;
        minimum = tokens.Count == 0 ? insets + 1D : tokens.Max(token => MeasureText(token, style.Font)) + insets;
        preferred = Math.Max(minimum, MeasureText(normalized, style.Font) + insets);
        if (style.ExplicitWidth.HasValue) {
            double authored = style.ExplicitWidth.Value + (style.BorderBox ? 0D : insets);
            minimum = Math.Max(minimum, authored);
            preferred = Math.Max(preferred, authored);
        }
    }

    private static IReadOnlyList<double> AllocateFixedColumnWidths(IReadOnlyList<double> requested, double totalWidth) {
        var result = requested.Select(value => Math.Max(0D, value)).ToArray();
        int unspecified = result.Count(value => value <= 0.0001D);
        double specifiedTotal = result.Sum();
        if (specifiedTotal > totalWidth && specifiedTotal > 0D) {
            double scale = totalWidth / specifiedTotal;
            for (int index = 0; index < result.Length; index++) result[index] = result[index] > 0D ? result[index] * scale : 0.01D;
        } else if (unspecified > 0) {
            double share = Math.Max(0.01D, (totalWidth - specifiedTotal) / unspecified);
            for (int index = 0; index < result.Length; index++) if (result[index] <= 0.0001D) result[index] = share;
        } else if (result.Length > 0) {
            double extra = (totalWidth - specifiedTotal) / result.Length;
            for (int index = 0; index < result.Length; index++) result[index] += extra;
        }
        NormalizeColumnWidthTotal(result, totalWidth);
        return result;
    }

    private static IReadOnlyList<double> AllocateAutoColumnWidths(
        IReadOnlyList<double> minimums,
        IReadOnlyList<double> preferred,
        double totalWidth) {
        var result = new double[minimums.Count];
        double minimumTotal = minimums.Sum();
        double preferredTotal = preferred.Sum();
        if (totalWidth <= minimumTotal + 0.0001D) {
            double scale = totalWidth / Math.Max(0.01D, minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = Math.Max(0.01D, minimums[index] * scale);
        } else if (totalWidth < preferredTotal - 0.0001D) {
            double progress = (totalWidth - minimumTotal) / Math.Max(0.01D, preferredTotal - minimumTotal);
            for (int index = 0; index < result.Length; index++) result[index] = minimums[index] + (preferred[index] - minimums[index]) * progress;
        } else {
            double extra = (totalWidth - preferredTotal) / result.Length;
            for (int index = 0; index < result.Length; index++) result[index] = preferred[index] + extra;
        }
        NormalizeColumnWidthTotal(result, totalWidth);
        return result;
    }

    private static void ApplySpanningWidth(double[] widths, int start, int span, double required) {
        double current = SumColumnWidths(widths, start, span);
        double deficit = required - current;
        if (deficit <= 0.0001D) return;
        double addition = deficit / span;
        for (int offset = 0; offset < span; offset++) widths[start + offset] += addition;
    }

    private static void NormalizeColumnWidthTotal(double[] widths, double totalWidth) {
        if (widths.Length == 0) return;
        widths[widths.Length - 1] += totalWidth - widths.Sum();
        widths[widths.Length - 1] = Math.Max(0.01D, widths[widths.Length - 1]);
    }

    private static double[] CreateColumnOffsets(IReadOnlyList<double> widths) {
        var offsets = new double[widths.Count];
        for (int index = 1; index < offsets.Length; index++) offsets[index] = offsets[index - 1] + widths[index - 1];
        return offsets;
    }

    private static double SumColumnWidths(IReadOnlyList<double> widths, int start, int count) {
        double result = 0D;
        for (int index = start; index < start + count && index < widths.Count; index++) result += widths[index];
        return result;
    }

    private static int DetermineDeclaredColumnCount(IElement table) {
        int count = 0;
        foreach (IElement element in table.QuerySelectorAll("col").Where(candidate => BelongsToTableColumn(candidate, table))) count += ReadSpan(element.GetAttribute("span"), 1000);
        return count;
    }

    private static bool BelongsToTableColumn(IElement column, IElement table) {
        IElement? current = column.ParentElement;
        while (current != null && !string.Equals(current.TagName, "table", StringComparison.OrdinalIgnoreCase)) current = current.ParentElement;
        return ReferenceEquals(current, table);
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
