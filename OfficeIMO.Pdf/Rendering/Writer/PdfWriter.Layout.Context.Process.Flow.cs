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
            if (flow.Options.MinimumRemainingHeight > 0D && available + 0.001D < flow.Options.MinimumRemainingHeight && y < GetCurrentFramePageStartY() - 0.001D) {
                NewPage();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                available = y - currentOpts.MarginBottom;
            }

            double? measuredHeight = MeasureFlowBlocks(blocks);
            double beforeFloatClearanceY = y;
            while (HasFloatingTables && (measuredHeight.GetValueOrDefault() > 0.001D || flow.Options.MinimumRemainingHeight > 0D) &&
                (flow.Options.KeepTogether || flow.Options.MinimumRemainingHeight > 0D ||
                 flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)) {
                double previousY = y;
                AvoidFloatingBlock(Math.Max(measuredHeight.GetValueOrDefault(), flow.Options.MinimumRemainingHeight));
                if (y >= previousY - 0.001D) break;
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }
                // Clearance changes page-top spacing and can change replayed content.
                // Recheck the complete group against any lower floating regions.
                measuredHeight = MeasureFlowBlocks(blocks);
                available = y - currentOpts.MarginBottom;
            }
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
            bool moveForMinimumHeight = flow.Options.MinimumRemainingHeight > 0D &&
                available + 0.001D < flow.Options.MinimumRemainingHeight;
            if ((moveForKeepTogether || moveForOverflow || moveForMinimumHeight) && y < GetCurrentFramePageStartY() - 0.001D) {
                NewPage();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                measuredHeight = MeasureFlowBlocks(blocks);
                fullPageMeasuredHeight = measuredHeight;
                available = y - currentOpts.MarginBottom;
                beforeFloatClearanceY = y;
                cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + 0.001D;
            }

            if (flow.Options.KeepTogether && fullPageMeasuredHeight.HasValue &&
                fullPageMeasuredHeight.Value > GetCurrentFramePageStartY() - currentOpts.MarginBottom + 0.001D) {
                throw new ArgumentException("Keep-together flow content exceeds the available full-page content height.");
            }

            if (cannotFitCurrentPage && flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.Skip) {
                // A skipped candidate does not reserve its temporary float clearance.
                y = beforeFloatClearanceY;
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
            var paintedRegions = new FloatingFlowCapture();
            if (capture != null) activeFloatingFlowCaptures.Push(paintedRegions);
            try { ProcessBlocks(blocks); }
            finally { if (capture != null) activeFloatingFlowCaptures.Pop(); }
            CaptureFlowRegions(capture, startPageNumber, startY, startOptions, paintedRegions);
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
            return GetCurrentFramePageStartY() - currentOpts.MarginBottom;
        }

        private double? MeasureFlowBlocks(IReadOnlyList<IPdfBlock> blocks) {
            return MeasureBlockSequence(blocks, currentOpts.MarginLeft, width, currentOpts.DefaultFontSize);
        }

        private void CaptureFlowRegions(PdfLayoutPositionCapture? capture, int startPageNumber, double startY, PdfOptions startOptions, FloatingFlowCapture paintedRegions) {
            if (capture == null) {
                return;
            }

            void AddCapturedRegion(PdfLayoutRegion region) {
                var painted = paintedRegions.PaintedRegions.Where(item => item.PageNumber == region.PageNumber).ToList();
                if (painted.Count == 0) { capture.Add(region); return; }
                // A float-only group leaves its flow cursor unchanged. Do not let that
                // nominal frame extend the painted table's captured bounds, even when
                // pagination synthesized a completed-page frame.
                if (paintedRegions.FlowPages.Contains(region.PageNumber) && region.Height > 0) painted.Add(region);
                double left = painted.Min(item => item.X), bottom = painted.Min(item => item.Y);
                double right = painted.Max(item => item.X + item.Width), top = painted.Max(item => item.Y + item.Height);
                capture.Add(new PdfLayoutRegion(region.PageNumber, left, bottom, right - left, top - bottom));
            }

            int endPageNumber = pages.Count + (currentPage == null ? 0 : 1);
            if (endPageNumber < startPageNumber) {
                capture.MarkSkipped();
                return;
            }

            if (endPageNumber == startPageNumber) {
                double bottom = Math.Min(startY, y);
                double height = Math.Max(0D, startY - y);
                AddCapturedRegion(new PdfLayoutRegion(startPageNumber, startOptions.MarginLeft, bottom, startOptions.PageWidth - startOptions.MarginLeft - startOptions.MarginRight, height));
                return;
            }

            AddCapturedRegion(new PdfLayoutRegion(
                startPageNumber,
                startOptions.MarginLeft,
                startOptions.MarginBottom,
                startOptions.PageWidth - startOptions.MarginLeft - startOptions.MarginRight,
                Math.Max(0D, startY - startOptions.MarginBottom)));

            for (int pageNumber = startPageNumber + 1; pageNumber < endPageNumber; pageNumber++) {
                LayoutResult.Page completedPage = pages[pageNumber - 1];
                PdfOptions options = completedPage.Options;
                AddCapturedRegion(new PdfLayoutRegion(
                    pageNumber,
                    options.MarginLeft,
                    options.MarginBottom,
                    options.PageWidth - options.MarginLeft - options.MarginRight,
                    options.PageHeight - options.MarginTop - options.MarginBottom));
            }

            if (currentPage != null) {
                AddCapturedRegion(new PdfLayoutRegion(
                    endPageNumber,
                    currentOpts.MarginLeft,
                    y,
                    currentOpts.PageWidth - currentOpts.MarginLeft - currentOpts.MarginRight,
                    Math.Max(0D, yStart - y)));
            }
        }
    }
}
