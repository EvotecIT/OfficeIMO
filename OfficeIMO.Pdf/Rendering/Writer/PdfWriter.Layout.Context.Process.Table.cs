using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderDeferredTableFlowBlock(DeferredTableBlock deferredTable, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex) {
            PdfTableStyle style = deferredTable.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            if (style.Position is { VerticalAlignment: not PdfTableVerticalAlignment.Top })
                throw new ArgumentException("Deferred floating tables support top alignment; use an eager table for center or bottom alignment.", nameof(deferredTable));
            double flowYBeforeTable = y;
            LayoutResult.Page? pageBeforeTable = currentPage;
            StringBuilder? pairedBorders = style.CellBorders?.Values.Any(HasPairedCellBorder) == true ? new StringBuilder() : null;
            foreach (DeferredTableBatch batch in deferredTable.CreateBatches(style)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (PdfFootnoteTraversal.References(batch.Table).Any())
                    throw new NotSupportedException("Page-associated footnote references are not supported in deferred tables. Use an eager table.");
                RenderTableFlowBlock(
                    batch.Table,
                    batch.IsLast ? nextBlock : null,
                    blockList,
                    blockIndex,
                    skipInitialHeaderRows: !batch.IsFirst,
                    bodyRowOffset: batch.BodyRowOffset,
                    logicalTopBoundary: batch.IsFirst,
                    logicalBottomBoundary: batch.IsLast,
                    restoreVerticalFlow: false,
                    preparedPairedStyle: pairedBorders == null ? null : PrepareDeferredPairedTableBorders(batch, style),
                    deferredPairedBorders: pairedBorders);
                if (!batch.IsLast) y -= GetTableCellSpacing(style);
            }
            if (pairedBorders != null) sb.Append(pairedBorders);
            if (style.Position != null || !style.ConsumesVerticalFlow && ReferenceEquals(currentPage, pageBeforeTable)) {
                y = ReferenceEquals(currentPage, pageBeforeTable) ? flowYBeforeTable : GetCurrentFramePageStartY();
            }
        }

        private void RenderTableFlowBlock(TableBlock tb, IPdfBlock? nextBlock, System.Collections.Generic.IList<IPdfBlock> blockList, int blockIndex, bool skipInitialHeaderRows = false, int bodyRowOffset = 0, bool logicalTopBoundary = true, bool logicalBottomBoundary = true, bool restoreVerticalFlow = true, PdfTableStyle? preparedPairedStyle = null, StringBuilder? deferredPairedBorders = null) {
            PdfTableStyle style = tb.Style ?? currentOpts.DefaultTableStyleSnapshot ?? TableStyles.Light();
            if (HasContainedCellBlocks(tb)) {
                if (skipInitialHeaderRows || bodyRowOffset != 0 || !logicalTopBoundary || !logicalBottomBoundary || !restoreVerticalFlow)
                    throw new NotSupportedException("Ordered cell content requires an eager inline table.");
                RenderContainedCellTable(tb, style, blockList, blockIndex, nextBlock);
                return;
            }
            double textClipBleed = style.ClipTextToCellBounds ? 0D : TableCellClipBleed;
            double flowYBeforeTable = y;
            LayoutResult.Page? pageBeforeTable = currentPage;
            int cols = GetTableColumnCount(tb);
            if (cols == 0) return;
            double padLeft = GetTableCellPaddingLeft(style);
            double padRight = GetTableCellPaddingRight(style);
            double padTop = GetTableCellPaddingTop(style);
            double padBottom = GetTableCellPaddingBottom(style);
            double cellSpacing = GetTableCellSpacing(style);
            double colGapPx = cellSpacing;
            double rowGapPx = cellSpacing;
            double size = GetTableBodyFontSize(style, currentOpts.DefaultFontSize);
            ValidateFlowTable(tb, style, cols);
            int headerRowCount = style.HeaderRowCount;
            int repeatHeaderRowCount = GetTableRepeatHeaderRowCount(style);
            int footerRowCount = style.FooterRowCount;
            int footerStartRowIndex = tb.Rows.Count - footerRowCount;

            style = preparedPairedStyle ?? PreparePairedTableBorders(tb, style);
            StringBuilder? pairedBorders = deferredPairedBorders;
            double contentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PreparedTableGeometry initialGeometry = PrepareTableGeometry(tb, style, contentWidth, logicalTopBoundary, logicalBottomBoundary);
            double tableWidth = initialGeometry.Width;
            double[] colPixel = initialGeometry.Widths;
            PreparedFlowTableRows preparedRows = initialGeometry.Rows;
            var rowLines = preparedRows.Lines;
            var rowLineCounts = preparedRows.LineCounts;
            var rowHeights = preparedRows.Heights;
            var rowLeadings = preparedRows.Leadings;
            var rowSizes = preparedRows.Sizes;
            var rowBold = preparedRows.Bold;
            var spanFlow = new TableSpanFlow(tb, style, cols, headerRowCount, footerStartRowIndex);
            bool tableHasFootnotes = ValidateTableFootnotePlacement(tb, style, cols);
            var remainingRowHeights = new double[rowHeights.Length + 1];
            for (int row = rowHeights.Length - 1; row >= 0; row--)
                remainingRowHeights[row] = rowHeights[row] +
                    GetTableRowGapAfter(row, rowHeights.Length, rowGapPx) + remainingRowHeights[row + 1];
            double tableCornerRadius = initialGeometry.CornerRadius;
            double xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);
            if (style.Position is { } horizontalPosition) xOrigin = PositionTableX(horizontalPosition, tableWidth) + style.HorizontalOffset;

            double TableBottom() => style.Position?.VerticalAnchor == PdfTableAnchor.Page ? 0 : FootnoteBodyBottom;
            double maxContentHeight = style.Position?.VerticalAnchor == PdfTableAnchor.Page ? currentOpts.PageHeight : GetFullPageContentHeight();
            bool tableFrameHasBodyContent = y < GetCurrentFramePageStartY() - .001D;
            if (style.Position is { VerticalAnchor: not PdfTableAnchor.Flow, VerticalAlignment: PdfTableVerticalAlignment.Top, VerticalOffset: > 0 } capacityPosition) {
                // Offset-preserving anchors cannot use space above their resolved top.
                maxContentHeight = Math.Min(maxContentHeight,
                    Math.Max(0, PositionTableY(capacityPosition, maxContentHeight) - TableBottom()));
            }
            string? captionText = string.IsNullOrWhiteSpace(style.Caption) ? null : style.Caption;
            System.Collections.Generic.IReadOnlyList<PdfTextRun>? captionRuns = null;
            System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>? captionLines = null;
            System.Collections.Generic.List<double>? captionLineHeights = null;
            double captionSize = style.CaptionFontSize ?? size;
            double captionLeading = captionSize * 1.25;
            double captionHeight = 0;
            if (captionText != null) {
                var captionFontForWrap = ChooseNormal(currentOpts.DefaultFont);
                captionRuns = new[] { PdfTextRun.Normal(captionText, style.CaptionColor, captionSize) };
                var captionWrap = WrapRichRunsCore(captionRuns, tableWidth, captionSize, captionFontForWrap, captionLeading, null, DefaultParagraphTabStopWidth, currentOpts);
                captionLines = captionWrap.Lines;
                captionLineHeights = captionWrap.LineHeights;
                captionHeight = MeasureRichLinesHeight(captionLineHeights, captionLines.Count, captionLeading);
            }

            double tableContentHeight = (captionLines == null ? 0 : captionHeight + style.CaptionSpacingAfter) + GetTableRowsHeight(rowHeights, 0, rowHeights.Length, rowGapPx);
            bool hasRepeatableHeader = repeatHeaderRowCount > 0 && tb.Rows.Count > headerRowCount;
            double repeatHeaderHeight = 0D;
            for (int header = 0; header < repeatHeaderRowCount; header++)
                repeatHeaderHeight += rowHeights[header] + GetTableRowGapAfter(header, tb.Rows.Count, rowGapPx);
            void ReflowTableForCurrentFrame(int rowIndex = 0, int startLine = 0, bool paddingChanged = false) {
                maxContentHeight = style.Position?.VerticalAnchor == PdfTableAnchor.Page ? currentOpts.PageHeight : GetFullPageContentHeight();
                double newContentWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                if (!paddingChanged && (activeColumnFlow == null || Math.Abs(contentWidth - newContentWidth) <= 0.001D)) return;
                contentWidth = newContentWidth;
                TableColumnLayout preparedColumns = ResolveTableColumnLayout(tb, currentOpts, style, cols, contentWidth, size, headerRowCount, footerStartRowIndex);
                tableWidth = preparedColumns.Width;
                colPixel = preparedColumns.Widths;
                preparedRows = PrepareFlowTableRows(tb, style, cols, colPixel, colGapPx, rowGapPx, headerRowCount, footerStartRowIndex,
                    preparedRows, rowIndex, startLine, spanFlow: spanFlow);
                rowLines = preparedRows.Lines;
                rowLineCounts = preparedRows.LineCounts;
                rowHeights = preparedRows.Heights;
                rowLeadings = preparedRows.Leadings;
                rowSizes = preparedRows.Sizes;
                rowBold = preparedRows.Bold;
                for (int row = rowHeights.Length - 1; row >= 0; row--)
                    remainingRowHeights[row] = rowHeights[row] + GetTableRowGapAfter(row, rowHeights.Length, rowGapPx) + remainingRowHeights[row + 1];
                tableCornerRadius = ResolveTableCornerRadius(style.CornerRadius, tableWidth,
                    logicalTopBoundary ? rowHeights[0] : double.PositiveInfinity,
                    logicalBottomBoundary ? rowHeights[rowHeights.Length - 1] : double.PositiveInfinity, colPixel[0], colPixel[colPixel.Length - 1]);
                if (captionRuns != null && rowIndex == 0 && startLine == 0) {
                    var wrap = WrapRichRunsCore(captionRuns, tableWidth, captionSize, ChooseNormal(currentOpts.DefaultFont), captionLeading,
                        null, DefaultParagraphTabStopWidth, currentOpts);
                    captionLines = wrap.Lines;
                    captionLineHeights = wrap.LineHeights;
                    captionHeight = MeasureRichLinesHeight(captionLineHeights, captionLines.Count, captionLeading);
                }
                tableContentHeight = (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter) + GetTableRowsHeight(rowHeights, 0, rowHeights.Length, rowGapPx);
            }
            if (logicalTopBoundary && style.Position is { } verticalPosition) {
                y = PositionTableY(verticalPosition, Math.Min(maxContentHeight, tableContentHeight));
                AvoidFloatingTable(verticalPosition, xOrigin, tableWidth, tableContentHeight);
            }
            void AdvanceInitialTablePage(bool preserveEmptyPage) {
                FlushPairedBorders();
                QueueColumnBalanceRemainder(tb, blockList, blockIndex);
                NewPage(preserveEmptyPage: preserveEmptyPage);
                tableFrameHasBodyContent = false;
                ReflowTableForCurrentFrame();
                xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);
                if (style.Position is { } position) {
                    flowYBeforeTable = y;
                    pageBeforeTable = currentPage;
                    xOrigin = PositionTableX(position, tableWidth) + style.HorizontalOffset;
                    y = PositionTableY(position, Math.Min(maxContentHeight, tableContentHeight));
                }
            }
            double tableSpacingBefore = style.Position == null && y < GetCurrentFramePageStartY() - 0.001 ? style.SpacingBefore : 0D;
            int initialAtomicFootnoteRows = 0;
            double initialAtomicFootnoteAfter = 0D;
            double MaximumContinuationHeight() => GetMaximumTableContinuationFrameHeight(style!, maxContentHeight);
            bool TryReserveTableFragment(int row, int start, int count, double height, bool commit, double extraBottom = 0D) =>
                !tableHasFootnotes && !HasPageFootnotes ||
                TryReserveFootnotes(tableHasFootnotes ? GetTableRowFootnoteSegments(tb, style, preparedRows, cols,
                    colPixel, colGapPx, rowGapPx, spanFlow, row, start, count, y, height, xOrigin,
                    Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), GetClosingTextPadding(style.SpacingAfter)) : Array.Empty<RichSeg>(),
                    y - height - extraBottom, commit);
            bool TryReserveTableGroup(int firstRow, int count, double groupTop, double bodyBottom, bool commit) =>
                !tableHasFootnotes && !HasPageFootnotes ||
                TryReserveFootnotes(tableHasFootnotes ? GetTableGroupFootnoteSegments(tb, style, preparedRows, cols,
                    colPixel, colGapPx, rowGapPx, spanFlow, firstRow, count, groupTop, xOrigin,
                    Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), GetClosingTextPadding(style.SpacingAfter)) : Array.Empty<RichSeg>(),
                    bodyBottom, commit);
            bool ProspectiveFragmentNeedsFootnotes(int row, int start, int count, double height) =>
                HasPageFootnotes || tableHasFootnotes && HasUnscheduledFootnote(GetTableRowFootnoteSegments(tb, style, preparedRows, cols,
                    colPixel, colGapPx, rowGapPx, spanFlow, row, start, count, y, height, xOrigin,
                    Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), GetClosingTextPadding(style.SpacingAfter)));
            void MovePastShortTableFrames(Func<double> requiredHeight, Action<bool> advance, bool searchRunningFrames = true) {
                RunningFrameAdmissionAttempt? runningAdmission = null;
                while (requiredHeight() > maxContentHeight + .001D) {
                    if (requiredHeight() <= MaximumContinuationHeight() + .001D &&
                        MaximumContinuationHeight() > maxContentHeight + .001D) {
                        advance(false);
                        continue;
                    }
                    if (!searchRunningFrames || style.Position != null || !TryAdvanceForRunningContent(ref runningAdmission,
                        requiredHeight(), advance, MaximumContinuationHeight())) break;
                }
            }
            double KeepHeight() => tableSpacingBefore + tableContentHeight + (style.Position == null ? style.SpacingAfter : 0D);
            void AdmitKeptTableFrame() {
                if (!style.KeepTogether) return;
                MovePastShortTableFrames(KeepHeight, preserve => { AdvanceInitialTablePage(preserve); tableSpacingBefore = 0D; });
                if (KeepHeight() > maxContentHeight + .001D)
                    throw new ArgumentException("Table height exceeds the available page content height.");
            }
            void NewInitialTablePage() {
                AdvanceInitialTablePage(preserveEmptyPage: false);
                tableSpacingBefore = 0D;
                // Every unpainted normal move must preserve the whole kept table,
                // even when KeepWithNext or another prepaint rule caused it.
                if (style.Position == null) AdmitKeptTableFrame();
            }
            if (style.KeepTogether) {
                AdmitKeptTableFrame();
                double keepHeight = KeepHeight();

                double GroupTop() => y - tableSpacingBefore - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                if ((y < GetCurrentFramePageStartY() - 0.001 || HasPageFootnotes) &&
                    (y - keepHeight < TableBottom() || !TryReserveTableGroup(0, tb.Rows.Count, GroupTop(), y - keepHeight, commit: false))) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                    keepHeight = KeepHeight();
                }
                while (!TryReserveTableGroup(0, tb.Rows.Count, GroupTop(), y - keepHeight, commit: false) && HasPageFootnotes) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                    keepHeight = KeepHeight();
                }
                if (!TryReserveTableGroup(0, tb.Rows.Count, GroupTop(), y - keepHeight, commit: false))
                    throw new ArgumentException("The kept table and its footnotes cannot fit within one page content frame.");
                initialAtomicFootnoteRows = tb.Rows.Count;
                initialAtomicFootnoteAfter = style.SpacingAfter;
            }

            if (style.KeepWithNext && nextBlock != null && style.Position == null) {
                double tableHeight = tableSpacingBefore + tableContentHeight + (style.Position == null ? style.SpacingAfter : 0D);
                bool restoresFlow = RestoresMeasuredTableFlow(style, tableHeight, y - TableBottom());
                double nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width,
                    currentOpts.DefaultFontSize, restoresFlow ? 0D : tableHeight);
                double keepHeight = restoresFlow ? Math.Max(tableHeight, nextHeight) : tableHeight + nextHeight;
                double groupTop = y - tableSpacingBefore - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                if (nextHeight > 0.001 && tableHeight <= maxContentHeight + 0.001 && keepHeight <= maxContentHeight + 0.001 &&
                    (y < GetCurrentFramePageStartY() - 0.001 || HasPageFootnotes) &&
                    (y - keepHeight < TableBottom() || !TryReserveTableGroup(0, tb.Rows.Count, groupTop, y - keepHeight, commit: false))) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                    tableHeight = tableContentHeight + style.SpacingAfter;
                    restoresFlow = RestoresMeasuredTableFlow(style, tableHeight, y - TableBottom());
                    nextHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex + 1, currentOpts.MarginLeft, width,
                        currentOpts.DefaultFontSize, restoresFlow ? 0D : tableHeight);
                    keepHeight = restoresFlow ? Math.Max(tableHeight, nextHeight) : tableHeight + nextHeight;
                    groupTop = y - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                }
                if (nextHeight > .001D && keepHeight <= y - TableBottom() + .001D &&
                    TryReserveTableGroup(0, tb.Rows.Count, groupTop, y - keepHeight, commit: false)) {
                    initialAtomicFootnoteRows = tb.Rows.Count;
                    initialAtomicFootnoteAfter = Math.Max(initialAtomicFootnoteAfter, style.SpacingAfter + nextHeight);
                }
            }

            if (style.KeepWithNext && nextBlock != null && style.Position != null) {
                double anchoredY = y;
                double keepHeight;
                try {
                    y = flowYBeforeTable;
                    keepHeight = MeasureKeepWithNextChainHeight(blockList, blockIndex, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize, 0);
                } finally { y = anchoredY; }
                if (keepHeight <= GetFullPageContentHeight() + 0.001 && flowYBeforeTable < GetCurrentFramePageStartY() - 0.001 &&
                    flowYBeforeTable - keepHeight < currentOpts.MarginBottom) NewInitialTablePage();
            }

            int minimumFirstPageBodyRows = Math.Min(
                style.MinimumBodyRowsOnFirstPage,
                Math.Max(0, footerStartRowIndex - headerRowCount));
            if (!skipInitialHeaderRows && minimumFirstPageBodyRows > 0) {
                int firstPageRowCount = headerRowCount + minimumFirstPageBodyRows;
                double FirstPageGroupHeight() =>
                    tableSpacingBefore +
                    (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter) +
                    GetTableRowsHeight(rowHeights, 0, firstPageRowCount, rowGapPx);
                double firstPageGroupHeight = FirstPageGroupHeight();
                if (firstPageGroupHeight <= MaximumContinuationHeight() + 0.001D &&
                    (y - firstPageGroupHeight < TableBottom() || !TryReserveTableGroup(0, firstPageRowCount,
                        y - tableSpacingBefore - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter),
                        y - firstPageGroupHeight, commit: false)) &&
                    (y < GetCurrentFramePageStartY() - 0.001D || HasPageFootnotes || MaximumContinuationHeight() > maxContentHeight + 0.001D)) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                    // This row-count preference is best effort; mandatory row/fragment
                    // admission below owns the running-variant search and its cycle.
                    MovePastShortTableFrames(FirstPageGroupHeight, AdvanceInitialTablePage, searchRunningFrames: false);
                    firstPageGroupHeight = FirstPageGroupHeight();
                }
                double groupTop = y - tableSpacingBefore - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                while (firstPageGroupHeight <= maxContentHeight + .001D && HasPageFootnotes &&
                    !TryReserveTableGroup(0, firstPageRowCount, groupTop, y - firstPageGroupHeight, commit: false)) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                    firstPageGroupHeight = FirstPageGroupHeight();
                    groupTop = y - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                }
                if (firstPageGroupHeight <= y - TableBottom() + .001D &&
                    TryReserveTableGroup(0, firstPageRowCount, groupTop, y - firstPageGroupHeight, commit: false))
                    initialAtomicFootnoteRows = Math.Max(initialAtomicFootnoteRows, firstPageRowCount);
            }

            if (tableSpacingBefore > 0) {
                if (y < GetCurrentFramePageStartY() - 0.001 && y - tableSpacingBefore < TableBottom()) {
                    NewInitialTablePage();
                    tableSpacingBefore = 0D;
                }

                y -= tableSpacingBefore;
            }

            if (initialAtomicFootnoteRows > 0) {
                double groupTop = y - (captionLines == null ? 0D : captionHeight + style.CaptionSpacingAfter);
                double bottom = groupTop - GetTableRowsHeight(rowHeights, 0, initialAtomicFootnoteRows, rowGapPx) -
                    initialAtomicFootnoteAfter;
                if (!TryReserveTableGroup(0, initialAtomicFootnoteRows, groupTop, bottom, commit: true))
                    throw new ArgumentException("The admitted table row group and its footnotes cannot fit within one page content frame.");
            }

            int? tableStructureElementIndex = null;
            LayoutResult.Page? tableStructurePage = null;
            int? EnsureTableStructureElement() {
                if (!emitGeneratedStructure || currentPage == null) {
                    return null;
                }

                if (!ReferenceEquals(tableStructurePage, currentPage)) {
                    tableStructurePage = currentPage;
                    tableStructureElementIndex = RegisterStructureContainer("Table", alternativeText: style!.AlternativeText);
                }

                return tableStructureElementIndex;
            }

            if (captionRuns != null && captionLines != null && captionLineHeights != null) {
                double CaptionAndFirstRowHeight() => captionHeight + style.CaptionSpacingAfter + (rowHeights.Length > 0 ? rowHeights[0] : 0D);
                MovePastShortTableFrames(CaptionAndFirstRowHeight, AdvanceInitialTablePage);
                if (CaptionAndFirstRowHeight() > maxContentHeight + 0.001D)
                    throw new ArgumentException("Table caption and first row exceed the available page content height.");
                var captionFont = ChooseNormal(currentOpts.DefaultFont);
                double firstRowHeight = rowHeights.Length > 0 ? rowHeights[0] : 0;
                double captionAndFirstRowHeight = captionHeight + style.CaptionSpacingAfter + firstRowHeight;
                if (y < GetCurrentFramePageStartY() - 0.001 &&
                    y - Math.Min(captionAndFirstRowHeight, maxContentHeight) < TableBottom()) {
                    NewInitialTablePage();
                }

                int? captionMarkedContentId = RegisterTextStructureElement("Caption", EnsureTableStructureElement());
                MarkRichFonts(captionRuns);
                if (style!.Position == null) RecordFlowPlacement(y);
                else ResolveFloatingBookmarks(y);
                WriteRichParagraph(sb, new RichParagraphBlock(captionRuns, style.CaptionAlign, style.CaptionColor), captionLines, captionLineHeights, currentOpts, FirstTextBaselineFromTop(captionFont, captionSize, y), captionSize, captionLeading, currentPage!.Annotations, xOrigin, tableWidth, structureType: "Caption", markedContentId: captionMarkedContentId, structurePage: currentPage, suppressInlineAccessibility: _suppressCanvasAccessibilityWrappers);
                if (style.Position is { } captionPosition)
                    ReserveFloatingTable(captionPosition, xOrigin, y, tableWidth, captionHeight + style.CaptionSpacingAfter);
                y -= captionHeight + style.CaptionSpacingAfter;
            }

            if (TableUsesBold(style, tb.Rows.Count, headerRowCount, footerStartRowIndex)) {
                currentPage!.UsedBold = true;
                usedBold = true;
            }

            bool ShouldBreakBefore(double rowHeight) =>
                (y < GetCurrentFramePageStartY() - 0.001 || HasPageFootnotes) &&
                y - rowHeight < TableBottom() &&
                rowHeight <= maxContentHeight;

            bool CanRepeatHeaderWithSegment(int rowIndex, int startLine, bool requireWholeRow) =>
                hasRepeatableHeader &&
                rowIndex >= headerRowCount &&
                repeatHeaderHeight + GetRequiredSegmentHeight(rowIndex, startLine, requireWholeRow) <= y - TableBottom() + 0.001;

            void ApplyTablePageContinuationSpacing(double requiredFirstSegmentHeight, int rowIndex, int startLine,
                bool requireWholeRow, bool continuingSpanTail) {
                double spacing = style.PageContinuationSpacingBefore;
                if (spacing <= 0D || y < GetCurrentFramePageStartY() - 0.001) {
                    return;
                }

                double available = y - TableBottom();
                if (tableHasFootnotes) {
                    double headers = hasRepeatableHeader && rowIndex >= headerRowCount ? repeatHeaderHeight : 0D;
                    int count = continuingSpanTail ? 0 : GetRequiredSegmentLineCount(rowIndex, startLine, requireWholeRow);
                    double fragmentHeight = Math.Max(0D, requiredFirstSegmentHeight - headers);
                    IEnumerable<RichSeg> references = GetTableRowFootnoteSegments(tb, style, preparedRows, cols,
                        colPixel, colGapPx, rowGapPx, spanFlow, rowIndex, startLine, count, y - headers, fragmentHeight,
                        xOrigin, Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), GetClosingTextPadding(style.SpacingAfter));
                    if (!TryReserveFootnotes(references, y - requiredFirstSegmentHeight, commit: false, out double reservedHeight))
                        return;
                    available -= Math.Max(0D, reservedHeight - (FootnoteBodyBottom - currentOpts.MarginBottom));
                }
                double spacingThatFits = Math.Min(spacing, Math.Max(0D, available - requiredFirstSegmentHeight));
                if (spacingThatFits > 0.001D) {
                    y -= spacingThatFits;
                }
            }

            int[] viewportRowGroups = GetTableViewportRowGroups(tb, cols);
            int GetRequiredSegmentLineCount(int rowIndex, int startLine, bool requireWholeRow) {
                int remaining = rowLineCounts[rowIndex] - startLine;
                if (requireWholeRow) return remaining;
                IReadOnlyList<TableCellLayout> cells = GetTableCellLayouts(tb, rowIndex, cols);
                for (int count = 1; count <= remaining; count++) {
                    if (IsTableRowFragmentBoundaryAllowed(rowLines[rowIndex], cells, startLine, count,
                        MaximumContinuationHeight(), requireDefaultFirstFragment: startLine == 0)) return count;
                }
                return remaining;
            }
            double GetRequiredSegmentHeight(int rowIndex, int startLine, bool requireWholeRow) {
                if (requireWholeRow) return GetTableViewportPlacementHeight(viewportRowGroups, rowHeights, rowIndex, rowGapPx);
                if (tableHasFootnotes || HasPageFootnotes) {
                    int legalCount = GetRequiredSegmentLineCount(rowIndex, startLine, false);
                    double legalHeight = MeasureTableRowSegmentHeight(rowIndex, startLine, legalCount, suppressCellObjects: false);
                    if (ProspectiveFragmentNeedsFootnotes(rowIndex, startLine, legalCount, legalHeight)) return legalHeight;
                }
                return MeasureTableRowSegmentHeight(rowIndex, startLine, 1, suppressCellObjects: false);
            }

            double GetTableContinuationRequiredHeight(int rowIndex, int startLine, bool requireWholeRow) {
                double rowRequiredHeight = GetRequiredSegmentHeight(rowIndex, startLine, requireWholeRow);
                if (hasRepeatableHeader && rowIndex >= headerRowCount) {
                    rowRequiredHeight += repeatHeaderHeight;
                }

                return rowRequiredHeight;
            }

            void DrawRepeatHeaders() {
                for (int headerIndex = 0; headerIndex < repeatHeaderRowCount; headerIndex++) {
                    DrawTableRow(headerIndex, renderAsHeader: true, suppressCellObjects: true);
                }
            }

            void FlushPairedBorders() {
                if (pairedBorders == null) return;
                sb.Append(pairedBorders);
                pairedBorders.Clear();
            }

            void NewTablePage(int rowIndex, int startLine = 0, bool requireWholeRow = false, bool continuingSpanTail = false,
                bool preserveEmptyPage = false, bool admitRunningFrame = true) {
                void AdvanceTableFrame(bool preserve) {
                    int? fragmentTable = spanFlow.ActiveCells.Any(span => span.Height > .001D) ? EnsureTableStructureElement() : tableStructureElementIndex;
                    FlushTableSpanFragments(tb, style, spanFlow, preparedRows, colPixel, colGapPx, rowGapPx, xOrigin, fragmentTable, ref pairedBorders,
                        cornerRadius: tableCornerRadius, fullFrameHeight: Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), logicalTopBoundary: logicalTopBoundary, logicalBottomBoundary: logicalBottomBoundary,
                        closingPadding: GetClosingTextPadding(style.SpacingAfter));
                    FlushPairedBorders();
                    PrepareTableFrameContinuation(tb, rowIndex);
                    if (CanQueueColumnBalanceRemainder(blockList)) {
                        ColumnFlowScope scope = activeColumnFlow!;
                        QueueColumnBalanceRemainder(MeasurePreparedTableColumnBalanceUnits(tb, style, preparedRows, cols, colPixel, rowGapPx,
                            MaximumContinuationHeight(), scope.Options.BalanceTableRowLines,
                            rowIndex, startLine), tb, blockList, blockIndex);
                    }
                    NewPage(preserveEmptyPage: preserve);
                    tableFrameHasBodyContent = false;
                    PdfTableStyle continuationStyle = PrepareTableBorderContinuation(tb, style, rowIndex);
                    bool paddingChanged = !ReferenceEquals(style, continuationStyle);
                    style = continuationStyle;
                    ReflowTableForCurrentFrame(rowIndex, startLine, paddingChanged);
                    ReserveTableSpanRemainder(spanFlow, style, preparedRows, colPixel, colGapPx, rowIndex, Math.Max(0D, maxContentHeight - (hasRepeatableHeader && rowIndex >= headerRowCount ? repeatHeaderHeight : 0D)));
                    repeatHeaderHeight = 0D;
                    for (int header = 0; header < repeatHeaderRowCount; header++)
                        repeatHeaderHeight += rowHeights[header] + GetTableRowGapAfter(header, tb.Rows.Count, rowGapPx);
                    xOrigin = ResolveTableX(tb.Align, style, currentOpts.MarginLeft, contentWidth, tableWidth);
                    if (style.Position is { } continuationPosition) {
                        flowYBeforeTable = y;
                        pageBeforeTable = currentPage;
                        xOrigin = PositionTableX(continuationPosition, tableWidth) + style.HorizontalOffset;
                        double continuationHeight = remainingRowHeights[rowIndex];
                        if (startLine > 0)
                            continuationHeight += MeasureTableRowSegmentHeight(rowIndex, startLine, rowLineCounts[rowIndex] - startLine, suppressCellObjects: false) - rowHeights[rowIndex];
                        if (hasRepeatableHeader && rowIndex >= headerRowCount) continuationHeight += repeatHeaderHeight;
                        y = PositionTableY(continuationPosition, Math.Min(maxContentHeight, continuationHeight));
                    }
                }
                double RequiredContinuationHeight() => continuingSpanTail
                        ? MeasureTableSpanRemainderHeight(spanFlow, style, preparedRows, colPixel, colGapPx, rowIndex,
                            minimumFragmentFrameHeight: Math.Max(0D, maxContentHeight - repeatHeaderHeight -
                                (rowIndex == tb.Rows.Count - 1 ? GetClosingTextPadding(style.SpacingAfter) : 0D))) +
                            (hasRepeatableHeader && rowIndex >= headerRowCount ? repeatHeaderHeight : 0D)
                        : GetTableContinuationRequiredHeight(rowIndex, startLine, requireWholeRow);
                double RequiredBodyFragmentHeight() {
                    if (continuingSpanTail)
                        return MeasureTableSpanRemainderHeight(spanFlow, style, preparedRows, colPixel, colGapPx, rowIndex,
                            minimumFragmentFrameHeight: Math.Max(0D, maxContentHeight -
                                (rowIndex == tb.Rows.Count - 1 ? GetClosingTextPadding(style.SpacingAfter) : 0D)));
                    bool wholeRowRequired = !GetTableRowAllowBreakAcrossPages(style, rowIndex) ||
                        viewportRowGroups[rowIndex] >= rowIndex;
                    return GetRequiredSegmentHeight(rowIndex, startLine, wholeRowRequired);
                }
                AdvanceTableFrame(preserveEmptyPage);
                // Admit only mandatory body content. Repeated headers and a
                // whole splittable row remain optional if this frame is shorter.
                // An outer row admission owns its own cycle search when present.
                if (admitRunningFrame) MovePastShortTableFrames(RequiredBodyFragmentHeight, AdvanceTableFrame);
                double requiredHeight = RequiredContinuationHeight();
                ApplyTablePageContinuationSpacing(requiredHeight, rowIndex, startLine, requireWholeRow, continuingSpanTail);
                if (continuingSpanTail ? hasRepeatableHeader && rowIndex >= headerRowCount && requiredHeight <= y - TableBottom() + .001D
                    : CanRepeatHeaderWithSegment(rowIndex, startLine, requireWholeRow)) {
                    DrawRepeatHeaders();
                }
            }

            double MeasureTableRowSegmentHeight(int rowIndex, int startLine, int lineCount, bool suppressCellObjects) =>
                MeasurePreparedTableRowSegmentHeight(tb, style, preparedRows, cols, colPixel, colGapPx,
                    rowIndex, startLine, lineCount, suppressCellObjects);

            int GetTableRowSegmentLineCountThatFits(int rowIndex, int startLine, double available,
                bool requireDefaultFirstFragment = false, bool? canMoveToNextFrame = null) {
                int take = GetPreparedTableRowSegmentLineCountThatFits(tb, style, preparedRows, cols, colPixel, colGapPx,
                    rowIndex, startLine, available, MaximumContinuationHeight(),
                    (canMoveToNextFrame ?? y < GetCurrentFramePageStartY() - 0.001D) ||
                        MaximumContinuationHeight() > maxContentHeight + 0.001D, requireDefaultFirstFragment);
                if (rowIndex == tb.Rows.Count - 1 && startLine + take == rowLineCounts[rowIndex] && GetClosingContainerPadding() > 0D)
                    take = GetPreparedTableRowSegmentLineCountThatFits(tb, style, preparedRows, cols, colPixel, colGapPx,
                        rowIndex, startLine, available - GetClosingTextPadding(style.SpacingAfter), MaximumContinuationHeight(),
                        (canMoveToNextFrame ?? y < GetCurrentFramePageStartY() - .001D) ||
                            MaximumContinuationHeight() > maxContentHeight + .001D, requireDefaultFirstFragment);
                while (take > 0) {
                    double height = MeasureTableRowSegmentHeight(rowIndex, startLine, take, suppressCellObjects: false);
                    double closing = rowIndex == tb.Rows.Count - 1 && startLine + take == rowLineCounts[rowIndex]
                        ? GetClosingTextPadding(style.SpacingAfter) : 0D;
                    if (TryReserveTableFragment(rowIndex, startLine, take, height, commit: false, closing)) break;
                    take = GetPreparedTableRowSegmentLineCountThatFits(tb, style, preparedRows, cols, colPixel, colGapPx,
                        rowIndex, startLine, available, MaximumContinuationHeight(),
                        (canMoveToNextFrame ?? y < GetCurrentFramePageStartY() - .001D) ||
                            MaximumContinuationHeight() > maxContentHeight + .001D, requireDefaultFirstFragment,
                        maximumLineCount: take - 1);
                }
                return take;
            }

            bool CanSplitTableRowIntoRemainingSpace(int rowIndex) =>
                rowIndex >= headerRowCount &&
                !TableRowHasViewport(tb, rowIndex, cols) &&
                GetTableRowAllowBreakAcrossPages(style, rowIndex) &&
                rowLineCounts[rowIndex] > 1 &&
                GetTableRowSegmentLineCountThatFits(rowIndex, 0, y - TableBottom(), requireDefaultFirstFragment: true) > 0;

            bool ShouldBreakBeforeFinalBodyRows(int rowIndex) {
                if (viewportRowGroups[rowIndex] >= rowIndex && !StartsTableViewportRowGroup(viewportRowGroups, rowIndex)) return false;
                int minimumBodyRows = Math.Min(style.MinimumBodyRowsOnLastPage, Math.Max(0, footerStartRowIndex - headerRowCount));
                if (minimumBodyRows <= 0 || footerStartRowIndex - rowIndex != minimumBodyRows) {
                    return false;
                }

                double currentRowHeight = rowHeights[rowIndex] + GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                double finalGroupHeight = remainingRowHeights[rowIndex];
                return ShouldBreakBeforeFinalTableBodyRows(
                    rowIndex,
                    headerRowCount,
                    footerStartRowIndex,
                    minimumBodyRows,
                    currentRowHeight,
                    finalGroupHeight,
                    y - TableBottom(),
                    hasRepeatableHeader ? repeatHeaderHeight : 0D,
                    maxContentHeight,
                    y < GetCurrentFramePageStartY() - 0.001);
            }

            void DrawTableRowSegment(int rowIndex, bool renderAsHeader, int startLine, int lineCount, bool suppressCellObjects = false, PageStructElement? existingRowStructureElement = null, double? continuedSpanHeight = null) {
                double rowHeight = continuedSpanHeight ?? MeasureTableRowSegmentHeight(rowIndex, startLine, lineCount, suppressCellObjects);
                double closing = rowIndex == tb.Rows.Count - 1 && startLine + lineCount == rowLineCounts[rowIndex]
                    ? GetClosingTextPadding(style.SpacingAfter) : 0D;
                if (!suppressCellObjects && (tableHasFootnotes || HasPageFootnotes) &&
                    !TryReserveTableFragment(rowIndex, startLine, lineCount, rowHeight, commit: true, closing))
                    throw new ArgumentException("The admitted table row fragment and its footnotes cannot fit within the page content frame.");
                if (!suppressCellObjects) tableFrameHasBodyContent = true;
                if (style.Position == null) RecordFlowPlacement(y);
                else ResolveFloatingBookmarks(y);
                var geometry = new PreparedTableGeometry(tb, style, preparedRows, colPixel, tableWidth, tableCornerRadius);
                var slice = new TableRowSlice(rowIndex, startLine, lineCount, rowHeight, renderAsHeader, suppressCellObjects);
                PaintPreparedTableRow(geometry, spanFlow, slice, xOrigin, y, EnsureTableStructureElement(),
                    ref pairedBorders, existingRowStructureElement, bodyRowOffset, logicalTopBoundary, logicalBottomBoundary);
                bool wholeRowSegment = startLine == 0 && lineCount == rowLineCounts[rowIndex];
                if (style.Position is { } floatingPosition) ReserveFloatingTable(floatingPosition, xOrigin, y, tableWidth, rowHeight + (wholeRowSegment ? GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx) : 0));
                if (wholeRowSegment || startLine + lineCount == rowLineCounts[rowIndex])
                    FlushTableSpanFragments(tb, style, spanFlow, preparedRows, colPixel, colGapPx, rowGapPx, xOrigin, EnsureTableStructureElement(), ref pairedBorders, rowIndex,
                        tableCornerRadius, Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), logicalTopBoundary, logicalBottomBoundary,
                        GetClosingTextPadding(style.SpacingAfter));
                RecordTableFrameContentBottom(tb, y - rowHeight);
                y -= rowHeight;
                if (wholeRowSegment && GetTableSpanRemainingLineCount(spanFlow, preparedRows, rowIndex) == 0)
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
            }

            void DrawTableRow(int rowIndex, bool renderAsHeader, bool suppressCellObjects = false) =>
                DrawTableRowSegment(rowIndex, renderAsHeader, 0, rowLineCounts[rowIndex], suppressCellObjects);

            bool TryDrawFootnoteLimitedSpanRow(int rowIndex) {
                if ((!tableHasFootnotes && !HasPageFootnotes) || !spanFlow.CoversRow(rowIndex) ||
                    GetTableRowFixedHeight(style, rowIndex).HasValue) return false;
                double maximum = Math.Min(rowHeights[rowIndex], y - TableBottom());
                double minimum = preparedRows.IntrinsicHeights[rowIndex];
                if (maximum < minimum - .001D) return false;
                if (!ProspectiveFragmentNeedsFootnotes(rowIndex, 0, rowLineCounts[rowIndex], maximum)) return false;
                foreach (double height in GetTableSpanFragmentHeightCandidates(spanFlow, style, preparedRows,
                    colPixel, colGapPx, rowGapPx, rowIndex, maximum, minimum)) {
                    double closing = rowIndex == tb.Rows.Count - 1 ? GetClosingTextPadding(style.SpacingAfter) : 0D;
                    if (!TryReserveTableFragment(rowIndex, 0, rowLineCounts[rowIndex], height, commit: false, closing)) continue;
                    DrawTableRowSegment(rowIndex, rowIndex < headerRowCount, 0, rowLineCounts[rowIndex], continuedSpanHeight: height);
                    if (ContinueCompletedRowSpans(rowIndex)) y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    return true;
                }
                return false;
            }

            void DrawSplitTableRow(int rowIndex, bool renderAsHeader) {
                int startLine = 0;
                int totalLines = rowLineCounts[rowIndex];
                PageStructElement? rowStructureElement = null;
                while (startLine < totalLines) {
                    MovePastShortTableFrames(() => MeasureTableRowSegmentHeight(rowIndex, startLine, 1, suppressCellObjects: false),
                        preserve => NewTablePage(rowIndex, startLine, preserveEmptyPage: preserve, admitRunningFrame: false));
                    totalLines = rowLineCounts[rowIndex];
                    double available = y - TableBottom();
                    double minimumRowSegmentHeight = MeasureTableRowSegmentHeight(rowIndex, startLine, 1, suppressCellObjects: false);
                    if (minimumRowSegmentHeight > maxContentHeight + 0.001D)
                        throw new ArgumentException("Table row content cannot fit within the available page content height.");
                    if (available < minimumRowSegmentHeight - 0.001) {
                        NewTablePage(rowIndex, startLine);
                        totalLines = rowLineCounts[rowIndex];
                        available = y - TableBottom();
                    }

                    int take = Math.Min(totalLines - startLine, GetTableRowSegmentLineCountThatFits(rowIndex, startLine, available));
                    while (take == 0) {
                        NewTablePage(rowIndex, startLine);
                        totalLines = rowLineCounts[rowIndex];
                        take = GetTableRowSegmentLineCountThatFits(rowIndex, startLine, y - TableBottom(), canMoveToNextFrame: false);
                        if (take == 0 && !HasPageFootnotes && MaximumContinuationHeight() <= maxContentHeight + 0.001D)
                            throw new ArgumentException("The table row reference and the first legal footnote fragment cannot fit within a fresh page content frame.");
                    }
                    DrawTableRowSegment(rowIndex, renderAsHeader && startLine == 0, startLine, take, existingRowStructureElement: rowStructureElement);
                    if (rowStructureElement == null && emitGeneratedStructure && currentPage != null) {
                        rowStructureElement = currentPage.StructElements.LastOrDefault(element => element.StructureType == "TR");
                    }
                    startLine += take;

                    if (startLine < totalLines) {
                        NewTablePage(rowIndex, startLine);
                        totalLines = rowLineCounts[rowIndex];
                    }
                }
            }

            bool ContinueCompletedRowSpans(int rowIndex) {
                if (GetTableRowFixedHeight(style, rowIndex).HasValue) return false;
                int remaining = GetTableSpanRemainingLineCount(spanFlow, preparedRows, rowIndex);
                bool continued = remaining > 0;
                while (remaining > 0) {
                    NewTablePage(rowIndex, rowLineCounts[rowIndex], continuingSpanTail: true);
                    // Reflow can increase the remaining wrapped-line count in a narrower column.
                    // Compare progress within the new frame's layout.
                    remaining = GetTableSpanRemainingLineCount(spanFlow, preparedRows, rowIndex);
                    if (remaining == 0) break;
                    double available = Math.Max(0D, y - TableBottom());
                    double required = MeasureTableSpanRemainderHeight(spanFlow, style, preparedRows, colPixel, colGapPx, rowIndex);
                    double closing = rowIndex == tb.Rows.Count - 1 ? GetClosingTextPadding(style.SpacingAfter) : 0D;
                    if (required <= available + .001D) available = Math.Max(0D, available - closing);
                    double height = Math.Min(required, available);
                    if (tableHasFootnotes || HasPageFootnotes) {
                        bool HasSpanProgress(double candidate) => spanFlow.ActiveCells.Any(span => span.EndRow == rowIndex + 1 &&
                            GetTableSpanAdmittedLineCount(span, style, preparedRows,
                                Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight),
                                span.EndRow == tb.Rows.Count ? closing : 0D, candidate, span.ProcessedHeight,
                                ResolveTableSpanAlignmentLeadingHeight(span, style, preparedRows, colPixel, colGapPx, rowGapPx)) > 0);
                        bool foundHeight = false;
                        foreach (double candidate in GetTableSpanFragmentHeightCandidates(spanFlow, style, preparedRows,
                            colPixel, colGapPx, rowGapPx, rowIndex, height, 0D)) {
                            if (!HasSpanProgress(candidate)) continue;
                            if (!TryReserveTableFragment(rowIndex, rowLineCounts[rowIndex], 0, candidate, commit: false, closing)) continue;
                            height = candidate;
                            foundHeight = true;
                            break;
                        }
                        if (!foundHeight) {
                            if (HasPageFootnotes) continue;
                            throw new ArgumentException("The merged cell reference and the first legal footnote fragment cannot fit within a fresh page content frame.");
                        }
                    }
                    DrawTableRowSegment(rowIndex, false, rowLineCounts[rowIndex], 0, continuedSpanHeight: height);
                    int next = GetTableSpanRemainingLineCount(spanFlow, preparedRows, rowIndex);
                    if (next >= remaining)
                        throw new ArgumentException("Merged table cell content cannot fit within the available continuation frame.");
                    remaining = next;
                }
                return continued;
            }

            int firstRowIndex = skipInitialHeaderRows ? headerRowCount : 0;
            for (int rowIndex = firstRowIndex; rowIndex < tb.Rows.Count; rowIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                ReserveTableSpanRemainder(spanFlow, style, preparedRows, colPixel, colGapPx, rowIndex, Math.Max(0D, maxContentHeight - (hasRepeatableHeader && rowIndex >= headerRowCount ? repeatHeaderHeight : 0D)));
                if (viewportRowGroups[rowIndex] >= rowIndex)
                    MovePastShortTableFrames(() => GetTableViewportPlacementHeight(viewportRowGroups, rowHeights, rowIndex, rowGapPx),
                        preserve => NewTablePage(rowIndex, requireWholeRow: true, preserveEmptyPage: preserve, admitRunningFrame: false));
                double placementHeight = GetTableViewportPlacementHeight(viewportRowGroups, rowHeights, rowIndex, rowGapPx);
                if (viewportRowGroups[rowIndex] >= rowIndex && placementHeight > maxContentHeight + 0.001D)
                    throw new ArgumentException("A cell viewport's complete visible row span must fit within one page; divide it into explicit fragments before rendering.");
                double requiredRowHeight = GetTableRowFixedHeight(style, rowIndex) ?? GetTableRowMinHeight(style, rowIndex);
                MovePastShortTableFrames(() => requiredRowHeight,
                    preserve => NewTablePage(rowIndex, preserveEmptyPage: preserve, admitRunningFrame: false));
                if (requiredRowHeight > maxContentHeight + 0.001D)
                    throw new ArgumentException("Table row height requirement exceeds the available page content height.");
                if (!GetTableRowAllowBreakAcrossPages(style, rowIndex))
                    MovePastShortTableFrames(() => rowHeights[rowIndex],
                        preserve => NewTablePage(rowIndex, requireWholeRow: true, preserveEmptyPage: preserve, admitRunningFrame: false));
                if (rowHeights[rowIndex] > maxContentHeight + 0.001) {
                    if (TableRowHasViewport(tb, rowIndex, cols))
                        throw new ArgumentException("A row containing a cell viewport must fit within one page; divide it into explicit fragments before rendering.");
                    if (!GetTableRowAllowBreakAcrossPages(style, rowIndex)) {
                        throw new ArgumentException("Table row height exceeds the available page content height and row splitting is disabled.");
                    }

                    DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    ContinueCompletedRowSpans(rowIndex);
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    continue;
                }

                double closingPadding = rowIndex == tb.Rows.Count - 1 ? GetClosingTextPadding(style.SpacingAfter) : 0D;
                if (ShouldBreakBefore(placementHeight + closingPadding) ||
                    !TryReserveTableFragment(rowIndex, 0, rowLineCounts[rowIndex], rowHeights[rowIndex], commit: false, closingPadding)) {
                    if (TryDrawFootnoteLimitedSpanRow(rowIndex)) continue;
                    if (CanSplitTableRowIntoRemainingSpace(rowIndex)) {
                        DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                        ContinueCompletedRowSpans(rowIndex);
                        y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                        continue;
                    }

                    NewTablePage(rowIndex, requireWholeRow: true);
                } else if (ShouldBreakBeforeFinalBodyRows(rowIndex)) {
                    NewTablePage(rowIndex, requireWholeRow: true);
                }

                placementHeight = GetTableViewportPlacementHeight(viewportRowGroups, rowHeights, rowIndex, rowGapPx);
                if (rowHeights[rowIndex] > maxContentHeight + 0.001D) {
                    if (TableRowHasViewport(tb, rowIndex, cols) || !GetTableRowAllowBreakAcrossPages(style, rowIndex))
                        throw new ArgumentException("Table row content cannot fit in the current column frame and row splitting is disabled.");
                    DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    ContinueCompletedRowSpans(rowIndex);
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    continue;
                }
                if (viewportRowGroups[rowIndex] >= rowIndex && placementHeight > y - TableBottom() + 0.001D)
                    throw new ArgumentException("A cell viewport's complete visible row span cannot fit in the current page frame.");

                int finalBodyMinimum = Math.Min(style.MinimumBodyRowsOnLastPage, Math.Max(0, footerStartRowIndex - headerRowCount));
                if (finalBodyMinimum > 0 && footerStartRowIndex - rowIndex == finalBodyMinimum) {
                    double GroupHeight() => remainingRowHeights[rowIndex] + GetClosingTextPadding(style.SpacingAfter);
                    bool GroupFits() => TryReserveTableGroup(rowIndex, tb.Rows.Count - rowIndex, y, y - GroupHeight(), commit: false);
                    if (GroupHeight() <= maxContentHeight + .001D && !GroupFits() &&
                        (tableFrameHasBodyContent || HasPageFootnotes))
                        NewTablePage(rowIndex, requireWholeRow: true);
                    while (GroupHeight() <= maxContentHeight + .001D && !GroupFits() && HasPageFootnotes)
                        NewTablePage(rowIndex, requireWholeRow: true);
                    if (GroupHeight() <= y - TableBottom() + .001D && GroupFits())
                        TryReserveTableGroup(rowIndex, tb.Rows.Count - rowIndex, y, y - GroupHeight(), commit: true);
                }

                bool footnoteLimitedRowDrawn = false;
                while (!TryReserveTableFragment(rowIndex, 0, rowLineCounts[rowIndex], rowHeights[rowIndex], commit: false, closingPadding)) {
                    if (TryDrawFootnoteLimitedSpanRow(rowIndex)) { footnoteLimitedRowDrawn = true; break; }
                    if (CanSplitTableRowIntoRemainingSpace(rowIndex)) break;
                    if (!HasPageFootnotes && !tableFrameHasBodyContent)
                        throw new ArgumentException("The table row reference and the first legal footnote fragment cannot fit within a fresh page content frame.");
                    NewTablePage(rowIndex, requireWholeRow: true);
                }
                if (footnoteLimitedRowDrawn) continue;
                if (!TryReserveTableFragment(rowIndex, 0, rowLineCounts[rowIndex], rowHeights[rowIndex], commit: false, closingPadding)) {
                    DrawSplitTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                    ContinueCompletedRowSpans(rowIndex);
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
                    continue;
                }

                DrawTableRow(rowIndex, renderAsHeader: rowIndex < headerRowCount);
                if (ContinueCompletedRowSpans(rowIndex))
                    y -= GetTableRowGapAfter(rowIndex, tb.Rows.Count, rowGapPx);
            }

            FlushTableSpanFragments(tb, style, spanFlow, preparedRows, colPixel, colGapPx, rowGapPx, xOrigin, EnsureTableStructureElement(), ref pairedBorders,
                cornerRadius: tableCornerRadius, fullFrameHeight: Math.Max(0D, MaximumContinuationHeight() - repeatHeaderHeight), logicalTopBoundary: logicalTopBoundary, logicalBottomBoundary: logicalBottomBoundary,
                closingPadding: GetClosingTextPadding(style.SpacingAfter));
            if (deferredPairedBorders == null) FlushPairedBorders();
            if (style.Position == null) y -= style.SpacingAfter;
            if (restoreVerticalFlow && (!style.ConsumesVerticalFlow || style.Position != null) && ReferenceEquals(currentPage, pageBeforeTable)) {
                y = flowYBeforeTable;
            }
        }

    }
}
