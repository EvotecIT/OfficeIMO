namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Flows one captured story at the current cursor using the shared legal-cut and table-fragment owner.</summary>
        private void RenderContinuedFlow(ContinuedFlowBlock story) {
            if (isRunningContent || activeColumnFlow != null || activeContainerScopes.Count > 0 || HasFloatingTables)
                throw new NotSupportedException("Continued stories require ordinary single-column flow without active floating tables or containers.");
            PdfDocument.ValidateFootnoteBlocks(story.Blocks, currentOpts);
            PdfDocument.ValidateFootnoteBlocks(story.Separator, currentOpts, allowTables: false);
            PdfDocument.ValidateFootnoteBlocks(story.ContinuationSeparator, currentOpts, allowTables: false);
            PdfDocument.ValidateFootnoteBlocks(story.ContinuationNotice, currentOpts, allowTables: false);
            var units = PrepareFootnoteUnits(story.Blocks);
            if (units.Count == 0) return;
            var first = PrepareFootnoteUnits(story.Separator);
            var later = PrepareFootnoteUnits(story.ContinuationSeparator);
            var notice = PrepareFootnoteUnits(story.ContinuationNotice);
            double noticeHeight = MeasureFootnoteUnits(notice);
            int start = 0;
            TableFlowCursor? table = null;
            bool continued = false;
            double? assignedMaximum = null;
            while (start < units.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<FootnoteUnit> separator = continued ? later : first;
                double separatorHeight = MeasureFootnoteUnits(separator);
                double separatorCredit = continued && story.AllowContinuationSeparatorInBottomMargin &&
                    notice.Count == 0 && !HasPageFootnotes && !HasSelectedContinuedFlowFooter()
                    ? Math.Min(separatorHeight, Math.Max(0D, FootnoteBodyBottom)) : 0D;
                double available = y - FootnoteBodyBottom - separatorHeight + separatorCredit;
                double maximum = GetCurrentFramePageStartY() - FootnoteBodyBottom - separatorHeight + separatorCredit - noticeHeight;
                if (assignedMaximum != maximum) {
                    foreach (FootnoteUnit unit in units) unit.MaximumKeepHeight = maximum;
                    assignedMaximum = maximum;
                }
                NoteFlowFragment? fragment = FitNoteFlowUnits(units, start, table,
                    available - (story.ReserveFinalNotice ? noticeHeight : 0D), clipTrailingSpacing: true);
                bool continues = fragment != null && fragment.NextStart < units.Count;
                if (continues && !story.ReserveFinalNotice && noticeHeight > 0D)
                    fragment = FitNoteFlowUnits(units, start, table, available - noticeHeight, clipTrailingSpacing: true);
                if (fragment == null) {
                    if (y >= GetCurrentFramePageStartY() - .001D && !HasPageFootnotes)
                        throw new ArgumentException("The first continued-story fragment and its separator or notice exceed the available page body height.");
                    NewPage();
                    continue;
                }
                continues = fragment.NextStart < units.Count;
                double top = y;
                PaintFootnoteUnits(separator, 0, separator.Count, ref top);
                PaintFootnoteUnits(units, fragment.Start, fragment.Count, ref top, fragment.Height,
                    fragment.Tables, fragment.StartTable != null);
                if (continues) PaintFootnoteUnits(notice, 0, notice.Count, ref top);
                y = top;
                start = fragment.NextStart;
                table = fragment.NextTable;
                if (!continues) break;
                continued = true;
                NewPage();
            }
        }

        // The derived frame can clear content after choosing its text fallback.
        // The original selected footer still owns that page's bottom margin.
        private bool HasSelectedContinuedFlowFooter() {
            PdfRunningContentContext context = currentPage?.RunningContentContext ?? GetRunningPageContext();
            int variant = currentPageBaseOptions.GetHeaderFooterVariantPageNumber(context.SectionPageNumber, context.PageNumber);
            return currentPageBaseOptions.GetRunningContentForPage(variant, isHeader: false) != null;
        }
    }
}
