namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Previews only references in the first legal visual fragment used by a preceding keep chain.</summary>
        private IEnumerable<RichSeg> GetFirstTableFootnoteSegments(TableBlock table) {
            PdfTableStyle style = table.Style ?? currentOpts.DefaultTableStyleSnapshot ?? PdfTableStyles.Light();
            int columns = GetTableColumnCount(table);
            if (columns == 0 || table.Rows.Count == 0 || !ValidateTableFootnotePlacement(table, style, columns))
                return Enumerable.Empty<RichSeg>();
            style = PreparePairedTableBorders(table, style);
            double gap = GetTableCellSpacing(style);
            int footerStart = table.Rows.Count - style.FooterRowCount;
            double fontSize = GetTableBodyFontSize(style, currentOpts.DefaultFontSize.ToPoints());
            TableColumnLayout columnLayout = ResolveTableColumnLayout(table, currentOpts, style, columns,
                currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints(),
                fontSize, style.HeaderRowCount, footerStart);
            PreparedFlowTableRows rows = PrepareFlowTableRows(table, style, columns, columnLayout.Widths,
                gap, gap, style.HeaderRowCount, footerStart);
            var spans = new TableSpanFlow(table, style, columns, style.HeaderRowCount, footerStart);
            double firstHeight = MeasureTableFirstFragmentHeight(table, style, rows, columns, columnLayout.Widths, gap, out _);
            int count = rows.LineCounts[0];
            if (!style.KeepTogether && GetTableRowAllowBreakAcrossPages(style, 0) && !TableRowHasViewport(table, 0, columns)) {
                IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(table, 0, columns);
                if (spans.CoversRow(0)) cells = cells.Where(cell => !spans.Contains(0, cell.Column)).ToArray();
                for (int candidate = 1; candidate <= rows.LineCounts[0]; candidate++) {
                    if (!IsTableRowFragmentBoundaryAllowed(rows.Lines[0], cells, 0, candidate,
                        GetMeasuredContinuationFrameHeight(), requireDefaultFirstFragment: true)) continue;
                    count = candidate;
                    break;
                }
            }
            return GetTableRowFootnoteSegments(table, style, rows, columns, columnLayout.Widths, gap, gap,
                spans, 0, 0, count, y, firstHeight, currentOpts.MarginLeft.ToPoints(),
                GetMeasuredContinuationFrameHeight(), GetClosingTextPadding(style.SpacingAfter.ToPoints()));
        }

        /// <summary>Preflights a breakable first row at its first legal cell boundary instead of requiring the entire row.</summary>
        private double MeasureTableFirstFragmentHeight(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
            int columns, double[] widths, double gap, out bool completesFirstRow) {
            double fullHeight = GetMeasuredContinuationFrameHeight();
            double headerGroup = MeasureTableHeaderFirstBodyHeight(style, rows.Heights, gap,
                row => MeasureTableFirstBodyPlacementHeight(table, style, rows, columns, widths, gap, row, fullHeight));
            if (headerGroup > 0D) {
                completesFirstRow = false;
                return headerGroup;
            }
            return MeasureTableFirstLegalRowFragmentHeight(table, style, rows, columns, widths, gap, 0, fullHeight, out completesFirstRow);
        }

        /// <summary>
        /// Measures what the first body row places beside the header: its first legal fragment, or the
        /// complete visible span when it starts a viewport row group, which must stay on one page.
        /// </summary>
        private static double MeasureTableFirstBodyPlacementHeight(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
            int columns, double[] widths, double gap, int row, double fullHeight) {
            int[] viewportGroups = GetTableViewportRowGroups(table, columns);
            return StartsTableViewportRowGroup(viewportGroups, row)
                ? GetTableViewportPlacementHeight(viewportGroups, rows.Heights, row, gap)
                : MeasureTableFirstLegalRowFragmentHeight(table, style, rows, columns, widths, gap, row, fullHeight, out _);
        }

        /// <summary>Measures a row's first legal fragment with the same boundary rules row placement uses.</summary>
        private static double MeasureTableFirstLegalRowFragmentHeight(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
            int columns, double[] widths, double gap, int row, double fullHeight, out bool completesRow) {
            completesRow = true;
            if (style.KeepTogether || !GetTableRowAllowBreakAcrossPages(style, row) || TableRowHasViewport(table, row, columns))
                return rows.Heights[row];
            IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(table, row, columns);
            var spanFlow = new TableSpanFlow(table, style, columns, style.HeaderRowCount,
                table.Rows.Count - style.FooterRowCount);
            bool hasIndependentSpan = spanFlow.CoversRow(row);
            if (hasIndependentSpan)
                cells = cells.Where(cell => !spanFlow.Contains(row, cell.Column)).ToArray();
            for (int count = 1; count <= rows.LineCounts[row]; count++) {
                if (IsTableRowFragmentBoundaryAllowed(rows.Lines[row], cells, 0, count, fullHeight, requireDefaultFirstFragment: true)) {
                    double height = MeasurePreparedTableRowSegmentHeight(table, style, rows, columns, widths, gap,
                        row, 0, count, suppressCellObjects: false);
                    completesRow = !hasIndependentSpan && count == rows.LineCounts[row];
                    return hasIndependentSpan ? Math.Min(height, rows.IntrinsicHeights[row]) : height;
                }
            }
            completesRow = !hasIndependentSpan;
            return rows.Heights[row];
        }

        /// <summary>Measures a row-column table row's first legal fragment, including the cell spacing a whole row reserves.</summary>
        private static double MeasureColTableFirstLegalRowFragmentHeight(ColTable table, int row, double fullHeight) {
            double gap = GetTableCellSpacing(table.Style);
            int[] viewportGroups = GetTableViewportRowGroups(table.Block, table.Columns);
            if (StartsTableViewportRowGroup(viewportGroups, row))
                return GetTableViewportPlacementHeight(viewportGroups, table.RowHeights, row, gap);
            int count = table.RowLineCounts[row];
            if (!table.Style.KeepTogether && GetTableRowAllowBreakAcrossPages(table.Style, row) &&
                !TableRowHasViewport(table.Block, row, table.Columns)) {
                IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(table.Block, row, table.Columns);
                for (int candidate = 1; candidate <= table.RowLineCounts[row]; candidate++) {
                    if (!IsTableRowFragmentBoundaryAllowed(table.RowLines[row], cells, 0, candidate, fullHeight,
                        requireDefaultFirstFragment: true)) continue;
                    count = candidate;
                    break;
                }
            }
            double height = MeasureColTableRowSegmentHeight(table, row, 0, count, suppressCellObjects: false);
            // A whole row reserves its trailing cell spacing when placed; a partial fragment does not.
            return height + (count == table.RowLineCounts[row] ? GetTableRowGapAfter(row, table.RowHeights.Length, gap) : 0D);
        }

        private static double MeasureColTableRowSegmentHeight(ColTable table, int rowIndex, int startLine, int lineCount,
            bool suppressCellObjects) {
            if (startLine == 0 && lineCount == table.RowLineCounts[rowIndex]) return table.RowHeights[rowIndex];
            TableBlock block = table.Block;
            PdfTableStyle style = table.Style;
            double columnGap = GetTableCellSpacing(style);
            double rowLeading = table.RowLeadings[rowIndex];
            double segmentHeight = GetTableRowInitialTextHeight(block, rowIndex, table.Columns, rowLeading) +
                GetTableRowMaxPaddingTop(block, style, rowIndex, table.Columns) +
                GetTableRowMaxPaddingBottom(block, style, rowIndex, table.Columns);
            foreach (TableCellLayout cell in GetTableCellLayouts(block, rowIndex, table.Columns)) {
                double cellWidth = GetTableCellWidth(table.ColumnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double innerWidth = cellWidth - GetTableCellPaddingLeft(style, rowIndex, cell.Column) -
                    GetTableCellPaddingRight(style, rowIndex, cell.Column);
                TableCellTextLayout lines = table.RowLines[rowIndex][cell.Column];
                int visibleLineCount = Math.Max(0, Math.Min(lineCount, lines.LineCount - startLine));
                double cellContentHeight = MeasureTableCellContentHeight(cell, lines, startLine, visibleLineCount, rowLeading, innerWidth,
                    !suppressCellObjects && startLine == 0) +
                    GetTableCellPaddingTop(style, rowIndex, cell.Column) + GetTableCellPaddingBottom(style, rowIndex, cell.Column);
                segmentHeight = Math.Max(segmentHeight, cellContentHeight);
            }

            // Keep first-fragment fitting consistent with the configured row height and drawing.
            return startLine == 0 ? Math.Max(segmentHeight, GetTableRowFixedHeight(style, rowIndex) ?? GetTableRowMinHeight(style, rowIndex)) : segmentHeight;
        }

        /// <summary>
        /// Measures the leading header rows together with the first legal fragment of the first body row.
        /// Returns zero when the style does not keep the header with that fragment or the table has no body row.
        /// </summary>
        private static double MeasureTableHeaderFirstBodyHeight(PdfTableStyle style, double[] heights, double gap,
            Func<int, double> measureFirstBodyFragment) {
            if (!KeepsRepeatedHeaderWithFirstBodyFragment(style, heights.Length)) return 0D;
            return GetTableRowsHeight(heights, 0, style.HeaderRowCount, gap) + measureFirstBodyFragment(style.HeaderRowCount);
        }

        /// <summary>Repeated header rows start with at least the first legal fragment of the following body row.</summary>
        private static bool KeepsRepeatedHeaderWithFirstBodyFragment(PdfTableStyle style, int rowCount) =>
            !style.KeepsRepeatedHeaderWithWholeRowsOnly && style.HeaderRowCount > 0 &&
            GetTableRepeatHeaderRowCount(style) > 0 && rowCount - style.FooterRowCount > style.HeaderRowCount;
    }
}
