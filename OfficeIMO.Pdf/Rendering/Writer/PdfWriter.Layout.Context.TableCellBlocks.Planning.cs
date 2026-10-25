namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Fits each cell independently to one shared physical row fragment.</summary>
        private TableRowSlice? PlanContainedCellRow(PreparedTableGeometry geometry, int row,
            IReadOnlyDictionary<int, CellFlowCursor>? positions, double available, double fullHeight,
            bool minimumOnly, out Dictionary<int, CellFlowCursor> next) {
            next = new Dictionary<int, CellFlowCursor>();
            var fragments = new Dictionary<int, NoteFlowFragment?>();
            double height = 0D;
            bool complete = true;
            foreach (PreparedCellFlow cell in geometry.CellFlows[row]) {
                CellFlowCursor cursor = positions != null && positions.TryGetValue(cell.Cell.Column, out var value) ? value : default;
                double pad = GetTableCellPaddingTop(geometry.Style, row, cell.Cell.Column) +
                    GetTableCellPaddingBottom(geometry.Style, row, cell.Cell.Column);
                if (cursor.Unit >= cell.Units.Count) {
                    fragments[cell.Cell.Column] = null;
                    next[cell.Cell.Column] = cursor;
                    continue;
                }
                foreach (FootnoteUnit unit in cell.Units) unit.MaximumKeepHeight = Math.Max(0D, fullHeight - pad);
                NoteFlowFragment? fragment = FitNoteFlowUnits(cell.Units, cursor.Unit, cursor.Table,
                    Math.Max(0D, available - pad), clipTrailingSpacing: false, minimumOnly, relaxOversizedParagraphKeeps: true);
                if (fragment == null) return null;
                bool cellComplete = fragment.NextStart >= cell.Units.Count;
                if (!cellComplete && (!GetTableRowAllowBreakAcrossPages(geometry.Style, row) ||
                    row < geometry.Style.HeaderRowCount && geometry.Rows.Heights[row] <= fullHeight + .001D)) return null;
                fragments[cell.Cell.Column] = fragment;
                next[cell.Cell.Column] = new CellFlowCursor(fragment.NextStart, fragment.NextTable);
                height = Math.Max(height, fragment.Height + pad);
                complete &= cellComplete;
            }
            if (positions == null) height = Math.Max(height, GetTableRowMinHeight(geometry.Style, row));
            if (height > available + .001D) return null;
            return new TableRowSlice(row, positions == null ? 0 : 1,
                complete ? geometry.Rows.LineCounts[row] : 0, height, row < geometry.Style.HeaderRowCount) {
                CellFragments = fragments, CompletesCellRow = complete
            };
        }

        /// <summary>Ordinary body tables use the same accepted nested-table plans as note stories.</summary>
        private void RenderContainedCellTable(TableBlock table, PdfTableStyle style, IList<IPdfBlock> blocks, int index, IPdfBlock? nextBlock) {
            if (isRunningContent || activeColumnFlow != null || HasFloatingTables || HasPictureExclusions)
                throw new NotSupportedException("Ordered table cell content requires an ordinary stable-width single-column frame.");
            double frameWidth = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
            PreparedTableGeometry geometry = PrepareTableGeometry(table, PreparePairedTableBorders(table, style), frameWidth);
            if (style.KeepWithNext && nextBlock != null) {
                double tableHeight = style.SpacingBefore + GetTableRowsHeight(geometry.Rows.Heights, 0, table.Rows.Count, geometry.Gap) + style.SpacingAfter;
                double next = MeasureKeepWithNextChainHeight(blocks, index + 1, currentOpts.MarginLeft, frameWidth,
                    currentOpts.DefaultFontSize, tableHeight);
                double kept = tableHeight + next + GetClosingContainerPadding();
                if (kept <= GetFullPageContentHeight() + .001D && kept > y - FootnoteBodyBottom + .001D)
                    NewPage();
            }
            TableFlowCursor? cursor = null;
            while (cursor == null || cursor.Row < table.Rows.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                double widthNow = currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight;
                if (Math.Abs(frameWidth - widthNow) > .001D)
                    throw new NotSupportedException("A continued ordered-cell table cannot change its containing frame width.");
                double closing = GetClosingContainerPadding();
                TableFragmentPlan? fragment = PlanTableFragment(geometry, cursor, y - FootnoteBodyBottom - closing, GetFullPageContentHeight() - closing);
                if (fragment == null) {
                    if (y >= GetCurrentFramePageStartY() - .001D)
                        throw new ArgumentException("The first legal nested-table row fragment exceeds the available page body height.");
                    PrepareTableFrameContinuation(table, cursor?.Row ?? 0);
                    NewPage();
                    continue;
                }
                double top = y;
                RecordFlowPlacement(top - fragment.Before);
                PaintTableFragment(fragment, 0D, frameWidth, ref top, noteContext: false);
                RecordTableFrameContentBottom(table, top + fragment.After + fragment.Slices[fragment.Slices.Count - 1].GapAfter);
                y = top;
                cursor = fragment.Next;
                if (!fragment.Continues) break;
                PrepareTableFrameContinuation(table, cursor.Row);
                NewPage();
            }
        }
    }
}
