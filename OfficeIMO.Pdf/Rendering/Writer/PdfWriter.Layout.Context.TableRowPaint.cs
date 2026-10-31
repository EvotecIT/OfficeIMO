namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Paints one accepted physical row slice. Page advancement and note admission stay with the caller.</summary>
        private void PaintPreparedTableRow(PreparedTableGeometry geometry, TableSpanFlow spanFlow,
            TableRowSlice slice, double xOrigin, double top, int? tableStructureElementIndex,
            ref StringBuilder? pairedBorders, PageStructElement? existingRowStructureElement = null,
            int bodyRowOffset = 0, bool logicalTopBoundary = true, bool logicalBottomBoundary = true,
            CollapsedTableBorders? collapsedBorders = null) {
            TableBlock tb = geometry.Table;
            PdfTableStyle style = geometry.Style;
            PreparedFlowTableRows preparedRows = geometry.Rows;
            int cols = geometry.Columns, headerRowCount = style.HeaderRowCount;
            int footerStartRowIndex = tb.Rows.Count - style.FooterRowCount;
            double[] colPixel = geometry.Widths;
            double colGapPx = geometry.Gap, rowGapPx = geometry.Gap, tableWidth = geometry.Width;
            double tableCornerRadius = geometry.CornerRadius, padTop = GetTableCellPaddingTop(style);
            var rowLines = preparedRows.Lines;
            var rowLineCounts = preparedRows.LineCounts;
            var rowHeights = preparedRows.Heights;
            var rowSizes = preparedRows.Sizes;
            var rowLeadings = preparedRows.Leadings;
            var rowBold = preparedRows.Bold;
            int rowIndex = slice.Row, startLine = slice.StartLine, lineCount = slice.LineCount;
            bool renderAsHeader = slice.IsHeader, suppressCellObjects = slice.IsRepeatedHeader;
            double y = top;
            bool renderAsFooter = rowIndex >= footerStartRowIndex;
            bool rowUsesBold = rowBold[rowIndex];
            double rowSize = rowSizes[rowIndex];
            double rowLeading = rowLeadings[rowIndex];
            if (rowUsesBold) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            var cells = GetTableCellLayouts(tb, rowIndex, cols);
            bool wholeRowSegment = slice.CellFragments != null || startLine == 0 && lineCount == rowLineCounts[rowIndex];
            double rowPadTop = GetTableRowMaxPaddingTop(tb, style, rowIndex, cols);
            double rowPadBottom = GetTableRowMaxPaddingBottom(tb, style, rowIndex, cols);
            double rowHeight = slice.Height;
            double rowBottom = y - rowHeight;
            if (!suppressCellObjects) spanFlow.AdmitRow(rowIndex, y, rowBottom, rowGapPx);
            if (currentOpts.Debug?.ShowTableRowBoxes == true) { pageDirty = true; DrawRowRect(sb, global::OfficeIMO.OfficeColor.FromNormalizedRgb(1, 0, 1), 0.6, xOrigin, rowBottom, tableWidth, rowHeight); }
            int bodyRowIndex = bodyRowOffset + rowIndex - headerRowCount;
            bool stripeBodyRow = bodyRowIndex >= 0 && bodyRowIndex % 2 == 1;
            bool[] rowFillSkips = GetRowSpanContinuationSkipColumns(tb, rowIndex, cols);
            double cornerRadius = tableCornerRadius;
            bool roundTop = cornerRadius > 0D && logicalTopBoundary && rowIndex == 0 && startLine == 0;
            bool roundBottom = cornerRadius > 0D && logicalBottomBoundary && rowIndex == tb.Rows.Count - 1 &&
                (slice.CellFragments != null ? slice.CompletesCellRow : startLine + lineCount >= rowLineCounts[rowIndex]);
            bool roundedRow = roundTop || roundBottom;
            if (roundedRow) { BeginRoundedClip(sb, xOrigin, rowBottom, tableWidth, rowHeight, cornerRadius, roundTop, roundTop, roundBottom, roundBottom); }
            if (style?.HeaderFill is not null && renderAsHeader) { pageDirty = true; DrawTableRowFill(sb, style.HeaderFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); } else if (style?.FooterFill is not null && renderAsFooter) { pageDirty = true; DrawTableRowFill(sb, style.FooterFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); } else if (!renderAsHeader && !renderAsFooter && style?.RowStripeFill is not null && stripeBodyRow) { pageDirty = true; DrawTableRowFill(sb, style.RowStripeFill.Value, xOrigin, colPixel, colGapPx, rowBottom, rowHeight, rowFillSkips, emitGeneratedStructure); }
            if (!renderAsHeader && !renderAsFooter && style?.BodyColumnFills != null) {
                bool[] bodyColumnFillSkips = GetMergedCellContinuationSkipColumns(tb, rowIndex, cols);
                double fillX = xOrigin;
                for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                    global::OfficeIMO.OfficeColor? fill = fillColumn < style.BodyColumnFills.Count ? style.BodyColumnFills[fillColumn] : null;
                    if (fill.HasValue && (fillColumn >= bodyColumnFillSkips.Length || !bodyColumnFillSkips[fillColumn])) {
                        pageDirty = true;
                        DrawRowFill(sb, fill.Value, fillX, rowBottom, colPixel[fillColumn], rowHeight, emitGeneratedStructure);
                    }
                    fillX += colPixel[fillColumn] + colGapPx;
                }
            }
            if (roundedRow) { EndRoundedClip(sb); }
            if (style?.CellFills != null && style.CellFills.Count > 0) {
                double fillX = xOrigin;
                for (int fillColumn = 0; fillColumn < cols; fillColumn++) {
                    if (!spanFlow.Contains(rowIndex, fillColumn) && style.CellFills.TryGetValue((rowIndex, fillColumn), out global::OfficeIMO.OfficeColor fill) &&
                        TryGetTableCellLayoutAtColumn(cells, fillColumn, out TableCellLayout fillCell) &&
                        (fillColumn >= rowFillSkips.Length || !rowFillSkips[fillColumn])) {
                        pageDirty = true;
                        int span = wholeRowSegment ? fillCell.ColumnSpan : 1;
                        double fillHeight = rowHeight;
                        double fillBottom = rowBottom;
                        if (wholeRowSegment) {
                            if (fillCell.RowSpan > 1) {
                                fillHeight = GetTableCellHeight(rowHeights, rowIndex, fillCell.RowSpan, rowGapPx);
                                fillBottom = y - fillHeight;
                            }
                        }

                        bool fillTouchesTop = cornerRadius > 0D && logicalTopBoundary && rowIndex == 0 && startLine == 0;
                        bool fillTouchesBottom = cornerRadius > 0D && logicalBottomBoundary &&
                            (wholeRowSegment
                                ? rowIndex + fillCell.RowSpan >= tb.Rows.Count
                                : rowIndex == tb.Rows.Count - 1 && startLine + lineCount >= rowLineCounts[rowIndex]);
                        bool fillTouchesLeft = fillColumn == 0;
                        bool fillTouchesRight = fillColumn + span >= cols;
                        bool clipCellFill = (fillTouchesTop || fillTouchesBottom) && (fillTouchesLeft || fillTouchesRight);
                        if (clipCellFill) {
                            BeginRoundedClip(sb, fillX, fillBottom, GetTableCellWidth(colPixel, fillColumn, span, colGapPx), fillHeight, cornerRadius,
                                fillTouchesTop && fillTouchesLeft,
                                fillTouchesTop && fillTouchesRight,
                                fillTouchesBottom && fillTouchesRight,
                                fillTouchesBottom && fillTouchesLeft);
                        }
                        DrawRowFill(sb, fill, fillX, fillBottom, GetTableCellWidth(colPixel, fillColumn, span, colGapPx), fillHeight, emitGeneratedStructure);
                        if (clipCellFill) { EndRoundedClip(sb); }
                    }
                    fillX += colPixel[fillColumn] + colGapPx;
                }
            }
            if (style != null && DrawTableCellDataBars(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                pageDirty = true;
            }
            if (style != null && DrawTableCellIcons(sb, style, cells, rowIndex, cols, xOrigin, y, rowBottom, rowHeight, colPixel, colGapPx, rowHeights, rowGapPx, wholeRowSegment, startLine, rowFillSkips, emitGeneratedStructure)) {
                pageDirty = true;
            }
            if (currentOpts.Debug?.ShowTableBaselines == true) {
                double x1 = xOrigin;
                double x2 = xOrigin + tableWidth;
                double baselineYDbg = y - padTop - GetAscenderForOptions(GetTableRowFont(currentOpts, rowUsesBold), rowSize, currentOpts);
                pageDirty = true;
                DrawHLine(sb, global::OfficeIMO.OfficeColor.FromNormalizedRgb(0, 0.6, 0), 0.4, x1, x2, baselineYDbg);
            }
            double xi = xOrigin;
            double yRect = rowBottom;
            double rowWidth = tableWidth;
            double hRect = rowHeight;
            if (style?.BorderColor is not null && style.BorderWidth.ToPoints() > 0) {
                pageDirty = true;
                bool[] topBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols);
                bool[] bottomBorderSkips = GetRowSpanBoundarySkipColumns(tb, rowIndex, cols);
                bool segmentBorderRows = HasSkippedColumns(topBorderSkips, cols) || HasSkippedColumns(bottomBorderSkips, cols);
                if (roundedRow) {
                    if (roundTop) {
                        DrawRoundedSideStroke(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), style.BorderWidth.ToPoints(), RoundedRectSide.Top, xOrigin, yRect, rowWidth, hRect, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    } else {
                        DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, colPixel, colGapPx, yRect + hRect, topBorderSkips, emitGeneratedStructure);
                    }
                    if (roundBottom) {
                        DrawRoundedSideStroke(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), style.BorderWidth.ToPoints(), RoundedRectSide.Bottom, xOrigin, yRect, rowWidth, hRect, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    } else {
                        DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, colPixel, colGapPx, yRect, bottomBorderSkips, emitGeneratedStructure);
                    }
                    DrawRoundedSideStroke(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), style.BorderWidth.ToPoints(), RoundedRectSide.Left, xOrigin, yRect, rowWidth, hRect, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                    DrawRoundedSideStroke(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), style.BorderWidth.ToPoints(), RoundedRectSide.Right, xOrigin, yRect, rowWidth, hRect, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                } else if (segmentBorderRows) {
                    DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, colPixel, colGapPx, yRect + hRect, topBorderSkips, emitGeneratedStructure);
                    DrawTableHorizontalLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, colPixel, colGapPx, yRect, bottomBorderSkips, emitGeneratedStructure);
                    DrawVLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, yRect + hRect, yRect, emitGeneratedStructure);
                    DrawVLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin + tableWidth, yRect + hRect, yRect, emitGeneratedStructure);
                } else {
                    DrawRoundedRowRect(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xOrigin, yRect, rowWidth, hRect, cornerRadius, roundTop, roundTop, roundBottom, roundBottom, emitGeneratedStructure);
                }

                double xi2 = xOrigin;
                double yTop = yRect + hRect;
                double yBottom = yRect;
                for (int c = 0; c < cols - 1; c++) {
                    xi2 += colPixel[c];
                    if (IsTableBoundaryInsideSpannedCell(tb, rowIndex, c, cols)) {
                        xi2 += colGapPx;
                        continue;
                    }

                    if (currentOpts.Debug?.ShowTableColumnGuides == true)
                        DrawVLine(sb, global::OfficeIMO.OfficeColor.FromNormalizedRgb(0, 0, 1), Math.Max(0.3, style.BorderWidth.ToPoints()), xi2, yTop, yBottom);
                    else
                        DrawVLine(sb, style.BorderColor.Value, style.BorderWidth.ToPoints(), xi2, yTop, yBottom, emitGeneratedStructure);
                    xi2 += colGapPx;
                }
            }
            if (style != null && renderAsFooter && rowIndex == footerStartRowIndex) {
                global::OfficeIMO.OfficeColor? footerSeparatorColor = style.FooterSeparatorColor ?? style.RowSeparatorColor;
                double footerSeparatorWidth = style.FooterSeparatorWidth.ToPoints() > 0 ? style.FooterSeparatorWidth.ToPoints() : style.RowSeparatorWidth.ToPoints();
                if (footerSeparatorColor is not null && footerSeparatorWidth > 0) {
                    pageDirty = true;
                    DrawTableHorizontalLine(sb, footerSeparatorColor.Value, footerSeparatorWidth, xOrigin, colPixel, colGapPx, y, GetRowSpanBoundarySkipColumns(tb, rowIndex - 1, cols), emitGeneratedStructure);
                }
            }
            global::OfficeIMO.OfficeColor? separatorColor = renderAsHeader && style?.HeaderSeparatorColor is not null ? style.HeaderSeparatorColor : style?.RowSeparatorColor;
            double separatorWidth = renderAsHeader && style?.HeaderSeparatorWidth.ToPoints() > 0 ? style.HeaderSeparatorWidth.ToPoints() : style?.RowSeparatorWidth.ToPoints() ?? 0;
            if (separatorColor is not null && separatorWidth > 0) {
                pageDirty = true;
                DrawTableHorizontalLine(sb, separatorColor.Value, separatorWidth, xOrigin, colPixel, colGapPx, rowBottom, GetRowSpanBoundarySkipColumns(tb, rowIndex, cols), emitGeneratedStructure);
            }
            var textColor = renderAsHeader ? style!.HeaderTextColor : renderAsFooter ? style!.FooterTextColor : style!.TextColor;
            int? rowStructureElementIndex = null;
            PageStructElement? rowStructureElement = existingRowStructureElement;
            if (rowStructureElement == null) {
                rowStructureElementIndex = RegisterStructureContainer("TR", tableStructureElementIndex);
                if (rowStructureElementIndex.HasValue) {
                    rowStructureElement = currentPage!.StructElements[rowStructureElementIndex.Value];
                }
            }
            if (!suppressCellObjects) spanFlow.BindStructureRow(currentPage, rowStructureElement);
            for (int cellIndex = 0; cellIndex < cells.Count; cellIndex++) {
                TableCellLayout cell = cells[cellIndex];
                if (spanFlow.Contains(rowIndex, cell.Column)) continue;
                int c = cell.Column;
                xi = xOrigin;
                for (int xColumn = 0; xColumn < c; xColumn++) {
                    xi += colPixel[xColumn] + colGapPx;
                }

                double cellWidth = GetTableCellWidth(colPixel, c, cell.ColumnSpan, colGapPx);
                double cellHeight = wholeRowSegment && cell.RowSpan > 1 ? GetTableCellHeight(rowHeights, rowIndex, cell.RowSpan, rowGapPx) : rowHeight;
                if (slice.CellFragments != null) {
                    PaintContainedCell(geometry, slice, cell, xi, y, cellWidth, cellHeight, rowStructureElementIndex, renderAsHeader);
                    continue;
                }
                TableCellTextLayout lines = rowLines[rowIndex][c];
                int sourceStartLine = cell.Viewport != null || wholeRowSegment && cell.RowSpan > 1 ? 0 : startLine;
                int requestedLineCount = cell.Viewport != null || wholeRowSegment && cell.RowSpan > 1 ? lines.LineCount : lineCount;
                RenderFlowTableCellContent(tb, style, cell, lines, rowIndex, sourceStartLine, requestedLineCount,
                    xi, y, cellWidth, cellHeight, rowSize, rowLeading, preparedRows.RunFontSizeScales[rowIndex],
                    rowUsesBold, renderAsHeader, wholeRowSegment, suppressCellObjects, textColor,
                    rowStructureElementIndex, rowStructureElement);
            }
            if (style?.CellBorders != null && style.CellBorders.Count > 0) {
                // On a rounded row a per-cell border (e.g. a left accent stripe or a top accent bar)
                // is drawn as a rounded-rectangle stroke clipped to its side + corners
                // (DrawRoundedCellBorderStrokes), so it shares the box's rounded corners and meets the
                // neighbouring border at the tangents. Elsewhere it stays the usual straight per-side
                // stroke (DrawCellBorder).
                double roundedOuterBorder = (style.BorderColor is not null && style.BorderWidth.ToPoints() > 0) ? style.BorderWidth.ToPoints() : 0D;
                double borderX = xOrigin;
                for (int borderColumn = 0; borderColumn < cols; borderColumn++) {
                    if (!spanFlow.Contains(rowIndex, borderColumn) && style.CellBorders.TryGetValue((rowIndex, borderColumn), out PdfCellBorder? cellBorder) &&
                        TryGetTableCellLayoutAtColumn(cells, borderColumn, out TableCellLayout borderCell) &&
                        (borderColumn >= rowFillSkips.Length || !rowFillSkips[borderColumn]) &&
                        HasCellBorderPaintOrConflict(cellBorder, collapsedBorders)) {
                        int span = wholeRowSegment || cellBorder.HasHiddenSegments ? borderCell.ColumnSpan : 1;
                        double borderHeight = hRect;
                        double borderBottom = yRect;
                        if (wholeRowSegment) {
                            if (borderCell.RowSpan > 1) {
                                borderHeight = GetTableCellHeight(rowHeights, rowIndex, borderCell.RowSpan, rowGapPx);
                                borderBottom = y - borderHeight;
                            }
                        }

                        pageDirty = true;
                        bool cellTouchesTop = cornerRadius > 0D && logicalTopBoundary && rowIndex == 0 && startLine == 0;
                        bool cellTouchesBottom = cornerRadius > 0D && logicalBottomBoundary &&
                            (wholeRowSegment
                                ? rowIndex + borderCell.RowSpan >= tb.Rows.Count
                                : rowIndex == tb.Rows.Count - 1 && startLine + lineCount >= rowLineCounts[rowIndex]);
                        bool cellTouchesLeft = borderColumn == 0;
                        bool cellTouchesRight = borderColumn + span >= cols;
                        bool topLeft = cellTouchesTop && cellTouchesLeft;
                        bool topRight = cellTouchesTop && cellTouchesRight;
                        bool bottomRight = cellTouchesBottom && cellTouchesRight;
                        bool bottomLeft = cellTouchesBottom && cellTouchesLeft;
                        StringBuilder borderOutput = HasPairedCellBorder(cellBorder) || spanFlow.HasDeferredFills
                            ? pairedBorders ??= new StringBuilder() : sb;
                        if (!cellBorder.HasHiddenSegments && (topLeft || topRight || bottomRight || bottomLeft)) {
                            DrawRoundedCellBorder(borderOutput, cellBorder, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight, cornerRadius, roundedOuterBorder, topLeft, topRight, bottomRight, bottomLeft, emitGeneratedStructure,
                                GetTableCellDiagonalFrame(borderCell, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight), collapsedBorders: collapsedBorders);
                        } else {
                            DrawCellBorder(borderOutput, cellBorder, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight, emitGeneratedStructure,
                                GetCellBorderSegmentLengths(rowHeights, rowIndex, borderCell.RowSpan, rowGapPx),
                                GetCellBorderSegmentLengths(colPixel, borderColumn, span, colGapPx),
                                GetTableCellDiagonalFrame(borderCell, borderX, borderBottom, GetTableCellWidth(colPixel, borderColumn, span, colGapPx), borderHeight), collapsedBorders: collapsedBorders);
                        }
                    }
                    borderX += colPixel[borderColumn] + colGapPx;
                }
            }
        }
    }
}
