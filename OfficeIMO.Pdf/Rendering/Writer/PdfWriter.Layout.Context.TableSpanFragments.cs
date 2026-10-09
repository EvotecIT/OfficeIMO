namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints merged-cell fragments only after their physical rows have been admitted to this frame.</summary>
        private void FlushTableSpanFragments(TableBlock table, PdfTableStyle style, TableSpanFlow flow,
            PreparedFlowTableRows rows, double[] columnWidths, double columnGap, double rowGap,
            double tableX, int? tableStructureElementIndex, ref StringBuilder? pairedBorders,
            int completedRow = -1, double cornerRadius = 0D, double fullFrameHeight = double.PositiveInfinity,
            bool logicalTopBoundary = true, bool logicalBottomBoundary = true, double closingPadding = 0D) {
            // Paint every deferred fill before any adjacent span's content or borders.
            foreach (TableSpanCellFlow span in flow.ActiveCells) {
                if (span.Height <= .001D || completedRow >= 0 && span.EndRow != completedRow + 1 ||
                    style.CellFills?.TryGetValue((span.Row, span.Cell.Column), out PdfColor fill) != true) continue;
                cancellationToken.ThrowIfCancellationRequested();
                TableCellLayout cell = span.Cell;
                double x = tableX;
                for (int column = 0; column < cell.Column; column++) x += columnWidths[column] + columnGap;
                double width = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                int admitted = GetTableSpanAdmittedLineCount(span, style, rows, fullFrameHeight,
                    span.EndRow == table.Rows.Count ? closingPadding : 0D);
                bool continues = completedRow < 0 || span.ConsumedLines + admitted < rows.Lines[span.Row][cell.Column].LineCount;
                var corners = GetTableSpanFragmentCorners(table, span, columnWidths.Length, cornerRadius,
                    logicalTopBoundary, logicalBottomBoundary, continues);
                bool rounded = corners.TopLeft || corners.TopRight || corners.BottomRight || corners.BottomLeft;
                if (rounded) BeginRoundedClip(sb, x, span.Top - span.Height, width, span.Height, cornerRadius,
                    corners.TopLeft, corners.TopRight, corners.BottomRight, corners.BottomLeft);
                DrawRowFill(sb, fill, x, span.Top - span.Height, width, span.Height, emitGeneratedStructure);
                if (rounded) EndRoundedClip(sb);
            }
            foreach (TableSpanCellFlow span in flow.ActiveCells) {
                if (span.Height <= .001D || completedRow >= 0 && span.EndRow != completedRow + 1) continue;
                cancellationToken.ThrowIfCancellationRequested();
                int row = span.Row;
                TableCellLayout cell = span.Cell;
                bool continuation = span.DestinationEmitted;
                double x = tableX;
                for (int column = 0; column < cell.Column; column++) x += columnWidths[column] + columnGap;
                double width = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap);
                double bottom = span.Top - span.Height;
                int? rowStructure = span.RowStructure == null ? RegisterStructureContainer("TR", tableStructureElementIndex) : null;
                TableCellTextLayout lines = rows.Lines[row][cell.Column];
                int admittedLines = GetTableSpanAdmittedLineCount(span, style, rows, fullFrameHeight,
                    span.EndRow == table.Rows.Count ? closingPadding : 0D);
                bool continues = completedRow < 0 || span.ConsumedLines + admittedLines < lines.LineCount;
                var corners = GetTableSpanFragmentCorners(table, span, columnWidths.Length, cornerRadius,
                    logicalTopBoundary, logicalBottomBoundary, continues);
                bool topLeft = corners.TopLeft, topRight = corners.TopRight;
                bool bottomLeft = corners.BottomLeft, bottomRight = corners.BottomRight;
                bool rounded = topLeft || topRight || bottomLeft || bottomRight;
                if (style.CellFills?.ContainsKey((row, cell.Column)) == true) {
                    // The opaque fill covered the earlier row grid. Restore its perimeter
                    // after all fills, before text and any explicit cell-border overrides.
                    if (style.BorderColor.HasValue && style.BorderWidth > 0D) {
                        var grid = new PdfCellBorder { Color = style.BorderColor, Width = style.BorderWidth };
                        if (rounded) DrawRoundedCellBorder(sb, grid, x, bottom, width, span.Height, cornerRadius,
                            style.BorderWidth, topLeft, topRight, bottomRight, bottomLeft, emitGeneratedStructure);
                        else DrawCellBorder(sb, grid, x, bottom, width, span.Height, emitGeneratedStructure);
                    }
                    if (span.FirstFragmentRow == row && row > 0) {
                        bool headerBefore = row - 1 < style.HeaderRowCount;
                        PdfColor? separator = headerBefore ? style.HeaderSeparatorColor ?? style.RowSeparatorColor : style.RowSeparatorColor;
                        double separatorWidth = headerBefore && style.HeaderSeparatorWidth > 0D ? style.HeaderSeparatorWidth : style.RowSeparatorWidth;
                        if (separator.HasValue && separatorWidth > 0D) DrawHLine(sb, separator.Value, separatorWidth, x, x + width, span.Top, emitGeneratedStructure);
                    }
                    if (!continues && style.RowSeparatorColor.HasValue && style.RowSeparatorWidth > 0D)
                        DrawHLine(sb, style.RowSeparatorColor.Value, style.RowSeparatorWidth, x, x + width, bottom, emitGeneratedStructure);
                }
                int take = RenderFlowTableCellContent(table, style, cell, lines, row, span.ConsumedLines,
                    admittedLines, x, span.Top, width, span.Height,
                    rows.Sizes[row], rows.Leadings[row], rows.RunFontSizeScales[row], rows.Bold[row],
                    renderAsHeader: false, wholeRowSegment: span.ConsumedLines == 0 && completedRow >= 0,
                    suppressCellObjects: true, style.TextColor, rowStructure, span.RowStructure,
                    fragmentRowSpan: span.LastFragmentRow - span.FirstFragmentRow + 1,
                    includeNamedDestination: !span.DestinationEmitted);
                span.ConsumedLines += take;
                span.DestinationEmitted = true;
                PdfCellBorder? border = null;
                bool hasBorderOverride = style.CellBorders?.TryGetValue((row, cell.Column), out border) == true;
                if (hasBorderOverride && HasRenderableCellBorder(border)) {
                    border = PrepareTableSpanFragmentBorder(border!, span, continuation, continues);
                    StringBuilder output = HasPairedCellBorder(border) || flow.HasDeferredFills(style)
                        ? pairedBorders ??= new StringBuilder() : sb;
                    if (rounded && !border.HasHiddenSegments)
                        DrawRoundedCellBorder(output, border, x, bottom, width, span.Height, cornerRadius,
                            style.BorderColor.HasValue ? style.BorderWidth : 0D, topLeft, topRight, bottomRight, bottomLeft,
                            emitGeneratedStructure, GetTableCellDiagonalFrame(cell, x, bottom, width, span.Height));
                    else
                        DrawCellBorder(output, border, x, bottom, width, span.Height, emitGeneratedStructure,
                            span.FragmentRowHeights.ToArray(),
                            GetCellBorderSegmentLengths(columnWidths, cell.Column, cell.ColumnSpan, columnGap),
                            GetTableCellDiagonalFrame(cell, x, bottom, width, span.Height));
                } else if (!hasBorderOverride && style.BorderColor.HasValue && style.BorderWidth > 0D && (continuation || continues)) {
                    // The ordinary row grid omits horizontal edges inside a span.
                    // Close that grid only where this span crosses a physical frame.
                    DrawCellBorder(sb, new PdfCellBorder {
                        Color = style.BorderColor, Width = style.BorderWidth,
                        Top = continuation, Bottom = continues, Left = false, Right = false
                    }, x, bottom, width, span.Height, emitGeneratedStructure);
                }
                span.Height = 0D;
            }
        }

        private static int GetTableSpanAdmittedLineCount(TableSpanCellFlow span, PdfTableStyle style,
            PreparedFlowTableRows rows, double fullFrameHeight, double closingPadding) {
            int row = span.Row;
            TableCellLayout cell = span.Cell;
            TableCellTextLayout lines = rows.Lines[row][cell.Column];
            double padding = GetTableCellPaddingTop(style, row, cell.Column) + GetTableCellPaddingBottom(style, row, cell.Column);
            double available = Math.Max(0D, span.Height - padding);
            int count = LimitTableCellLineCountToHeight(lines, span.ConsumedLines,
                Math.Max(0, lines.LineCount - span.ConsumedLines), rows.Leadings[row], available);
            double fullTextHeight = Math.Max(0D, fullFrameHeight - closingPadding - padding);
            return LimitTableRowFragmentToParagraphBoundaries(rows.Lines[row], new[] { cell },
                span.ConsumedLines, count, fullTextHeight, available + .001D < fullTextHeight);
        }

        private static (bool TopLeft, bool TopRight, bool BottomRight, bool BottomLeft) GetTableSpanFragmentCorners(
            TableBlock table, TableSpanCellFlow span, int columns, double radius,
            bool logicalTop, bool logicalBottom, bool continues) {
            bool top = radius > 0D && logicalTop && span.Row == 0 && !span.DestinationEmitted;
            bool bottom = radius > 0D && logicalBottom && span.EndRow == table.Rows.Count && !continues;
            bool left = span.Cell.Column == 0, right = span.Cell.Column + span.Cell.ColumnSpan == columns;
            return (top && left, top && right, bottom && right, bottom && left);
        }

        /// <summary>Reserves any text carried from earlier frames in a flexible span's final physical row.</summary>
        private static void ReserveTableSpanRemainder(TableSpanFlow flow, PdfTableStyle style,
            PreparedFlowTableRows rows, double[] columnWidths, double columnGap, int rowIndex, double maximumHeight) {
            if (GetTableRowFixedHeight(style, rowIndex).HasValue) return;
            if (flow.CoversRow(rowIndex))
                rows.Heights[rowIndex] = Math.Max(rows.IntrinsicHeights[rowIndex], Math.Min(rows.Heights[rowIndex], maximumHeight));
            foreach (TableSpanCellFlow span in flow.ActiveCells) {
                if (span.EndRow != rowIndex + 1) continue;
                TableCellLayout cell = span.Cell;
                TableCellTextLayout lines = rows.Lines[span.Row][cell.Column];
                int remaining = Math.Max(0, lines.LineCount - span.ConsumedLines);
                if (remaining == 0) continue;
                double innerWidth = GetTableCellWidth(columnWidths, cell.Column, cell.ColumnSpan, columnGap) -
                    GetTableCellPaddingLeft(style, span.Row, cell.Column) - GetTableCellPaddingRight(style, span.Row, cell.Column);
                double required = MeasureTableCellContentHeight(cell, lines, span.ConsumedLines, remaining,
                    rows.Leadings[span.Row], innerWidth, includeObjects: false) +
                    GetTableCellPaddingTop(style, span.Row, cell.Column) + GetTableCellPaddingBottom(style, span.Row, cell.Column);
                rows.Heights[rowIndex] = Math.Max(rows.Heights[rowIndex], Math.Min(required - span.Height, maximumHeight));
            }
        }

        private static int GetTableSpanRemainingLineCount(TableSpanFlow flow, PreparedFlowTableRows rows, int rowIndex) =>
            flow.ActiveCells.Where(span => span.EndRow == rowIndex + 1)
                .Select(span => Math.Max(0, rows.Lines[span.Row][span.Cell.Column].LineCount - span.ConsumedLines)).DefaultIfEmpty(0).Max();

        private static double MeasureTableSpanRemainderHeight(TableSpanFlow flow, PdfTableStyle style,
            PreparedFlowTableRows rows, double[] widths, double columnGap, int rowIndex,
            double? minimumFragmentFrameHeight = null, bool subtractAdmittedHeight = false) {
            double height = 0D;
            foreach (TableSpanCellFlow span in flow.ActiveCells) {
                if (span.EndRow != rowIndex + 1) continue;
                TableCellTextLayout lines = rows.Lines[span.Row][span.Cell.Column];
                int count = Math.Max(0, lines.LineCount - span.ConsumedLines);
                if (count == 0) continue;
                double innerWidth = GetTableCellWidth(widths, span.Cell.Column, span.Cell.ColumnSpan, columnGap) -
                    GetTableCellPaddingLeft(style, span.Row, span.Cell.Column) - GetTableCellPaddingRight(style, span.Row, span.Cell.Column);
                if (minimumFragmentFrameHeight.HasValue) {
                    double textHeight = Math.Max(0D, minimumFragmentFrameHeight.Value -
                        GetTableCellPaddingTop(style, span.Row, span.Cell.Column) - GetTableCellPaddingBottom(style, span.Row, span.Cell.Column));
                    int maximum = LimitTableCellLineCountToHeight(lines, span.ConsumedLines, count, rows.Leadings[span.Row], textHeight);
                    // Use the same keep/widow boundary as admission. If no boundary fits
                    // a fresh frame, retain its existing relaxation to a single line.
                    count = 1;
                    for (int candidate = 1; candidate <= maximum; candidate++) {
                        if (!IsTableRowFragmentBoundaryAllowed(rows.Lines[span.Row], new[] { span.Cell },
                            span.ConsumedLines, candidate, textHeight)) continue;
                        count = candidate;
                        break;
                    }
                }
                double required = MeasureTableCellContentHeight(span.Cell, lines, span.ConsumedLines,
                    count, rows.Leadings[span.Row], innerWidth, includeObjects: false) +
                    GetTableCellPaddingTop(style, span.Row, span.Cell.Column) + GetTableCellPaddingBottom(style, span.Row, span.Cell.Column);
                height = Math.Max(height, required - (subtractAdmittedHeight ? span.Height : 0D));
            }
            return height;
        }
    }
}
