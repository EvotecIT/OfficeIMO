namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private void RenderFlowBlock(FlowBlock flow) {
            PdfLayoutPositionCapture? capture = flow.Capture;
            if (capture != null && initializedPositionCaptures.Add(capture)) {
                capture.BeginLayoutPass();
            }

            PdfFlowContext context = CreateFlowContext();
            if (flow.Options.ShowIf != null && !flow.Options.ShowIf(context)) {
                capture?.MarkSkipped();
                return;
            }

            IReadOnlyList<IPdfBlock> blocks = MaterializeFlow(flow, context);
            double available = y - currentOpts.MarginBottom;
            if (flow.Options.MinimumRemainingHeight > 0D && available + 0.001D < flow.Options.MinimumRemainingHeight && y < yStart - 0.001D) {
                NewPage();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                available = y - currentOpts.MarginBottom;
            }

            double? measuredHeight = MeasureFlowBlocks(blocks);
            if (!measuredHeight.HasValue &&
                (flow.Options.KeepTogether || flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)) {
                throw new NotSupportedException("KeepTogether and non-continuing overflow behavior require flow content whose height can be determined before rendering. Remove the constraint or move dynamic, multi-column, deferred-table, table-of-contents, canvas, or explicit page-boundary content outside the flow.");
            }

            bool cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + 0.001D;
            double? fullPageMeasuredHeight = measuredHeight.HasValue
                ? MeasureBlockSequenceAtFrameStart(blocks, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize)
                : null;
            bool fitsFullPage = fullPageMeasuredHeight.HasValue &&
                                fullPageMeasuredHeight.Value <= GetCurrentFramePageStartY() - currentOpts.MarginBottom + 0.001D;
            bool moveForKeepTogether = flow.Options.KeepTogether && cannotFitCurrentPage && fitsFullPage;
            bool moveForOverflow = flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.MoveToNextPage && cannotFitCurrentPage && fitsFullPage;
            if ((moveForKeepTogether || moveForOverflow) && y < yStart - 0.001D) {
                NewPage();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                measuredHeight = MeasureFlowBlocks(blocks);
                fullPageMeasuredHeight = measuredHeight;
                available = y - currentOpts.MarginBottom;
                cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + 0.001D;
            }

            if (flow.Options.KeepTogether && fullPageMeasuredHeight.HasValue &&
                fullPageMeasuredHeight.Value > GetCurrentFramePageStartY() - currentOpts.MarginBottom + 0.001D) {
                throw new ArgumentException("Keep-together flow content exceeds the available full-page content height.");
            }

            if (cannotFitCurrentPage && flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Skip) {
                capture?.MarkSkipped();
                return;
            }

            if (cannotFitCurrentPage && flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.StopDocument) {
                capture?.MarkSkipped();
                stopDocumentFlow = true;
                return;
            }

            int startPageNumber = pages.Count + 1;
            double startY = y;
            PdfOptions startOptions = currentOpts;
            ProcessBlocks(blocks);
            CaptureFlowRegions(capture, startPageNumber, startY, startOptions);
        }

        private IReadOnlyList<IPdfBlock> MaterializeFlow(FlowBlock flow, PdfFlowContext context) {
            if (!flow.IsReplayable) return flow.Materialize(context);
            var key = new FlowMaterializationKey(flow, context);
            if (!deferredMaterializations.TryGetValue(key, out IReadOnlyList<IPdfBlock>? blocks)) {
                blocks = flow.Materialize(context);
                deferredMaterializations.Add(key, blocks);
            }

            return blocks;
        }

        private PdfFlowContext CreateFlowContext() {
            return new PdfFlowContext(
                pages.Count + 1,
                y - currentOpts.MarginBottom,
                GetFullPageContentHeight(),
                width,
                currentOpts.PageWidth,
                currentOpts.PageHeight);
        }

        private double GetFullPageContentHeight() {
            return currentOpts.PageHeight - currentOpts.MarginTop - currentOpts.MarginBottom;
        }

        private double? MeasureFlowBlocks(IReadOnlyList<IPdfBlock> blocks) {
            return MeasureBlockSequence(blocks, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
        }

        private void CaptureFlowRegions(PdfLayoutPositionCapture? capture, int startPageNumber, double startY, PdfOptions startOptions) {
            if (capture == null) {
                return;
            }

            int endPageNumber = pages.Count + (currentPage == null ? 0 : 1);
            if (endPageNumber < startPageNumber) {
                capture.MarkSkipped();
                return;
            }

            if (endPageNumber == startPageNumber) {
                double bottom = Math.Min(startY, y);
                double height = Math.Max(0D, startY - y);
                capture.Add(new PdfLayoutRegion(startPageNumber, startOptions.MarginLeft, bottom, startOptions.PageWidth - startOptions.MarginLeft - startOptions.MarginRight, height));
                return;
            }

            capture.Add(new PdfLayoutRegion(
                startPageNumber,
                startOptions.MarginLeft,
                startOptions.MarginBottom,
                startOptions.PageWidth - startOptions.MarginLeft - startOptions.MarginRight,
                Math.Max(0D, startY - startOptions.MarginBottom)));

            for (int pageNumber = startPageNumber + 1; pageNumber < endPageNumber; pageNumber++) {
                LayoutResult.Page completedPage = pages[pageNumber - 1];
                PdfOptions options = completedPage.Options;
                capture.Add(new PdfLayoutRegion(
                    pageNumber,
                    options.MarginLeft,
                    options.MarginBottom,
                    options.PageWidth - options.MarginLeft - options.MarginRight,
                    options.PageHeight - options.MarginTop - options.MarginBottom));
            }

            if (currentPage != null) {
                capture.Add(new PdfLayoutRegion(
                    endPageNumber,
                    currentOpts.MarginLeft,
                    y,
                    currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight,
                    Math.Max(0D, yStart - y)));
            }
        }
    }
}
