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
            double maximum = GetFullPageContentHeight() - Math.Max(MeasureFootnoteUnits(first), MeasureFootnoteUnits(later)) - noticeHeight;
            foreach (FootnoteUnit unit in units) unit.MaximumKeepHeight = maximum;
            int start = 0;
            TableFlowCursor? table = null;
            bool continued = false;
            while (start < units.Count) {
                cancellationToken.ThrowIfCancellationRequested();
                IReadOnlyList<FootnoteUnit> separator = continued ? later : first;
                double separatorHeight = MeasureFootnoteUnits(separator);
                double available = y - FootnoteBodyBottom - separatorHeight;
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
    }
}
