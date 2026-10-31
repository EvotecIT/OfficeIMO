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
            double available = y - FootnoteBodyBottom;
            if (flow.Options.MinimumRemainingHeight.ToPoints() > 0D && available + 0.001D < flow.Options.MinimumRemainingHeight.ToPoints() && y < GetCurrentFramePageStartY() - 0.001D) {
                NewBlockFrame();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                available = y - FootnoteBodyBottom;
            }

            double? measuredHeight = MeasureFlowBlocks(blocks);
            double beforeFloatClearanceY = y;
            while (HasFloatingTables && (measuredHeight.GetValueOrDefault() > 0.001D || flow.Options.MinimumRemainingHeight.ToPoints() > 0D) &&
                (flow.Options.KeepTogether || flow.Options.MinimumRemainingHeight.ToPoints() > 0D ||
                 flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)) {
                double previousY = y;
                AvoidFloatingBlock(Math.Max(measuredHeight.GetValueOrDefault(), flow.Options.MinimumRemainingHeight.ToPoints()));
                if (y >= previousY - 0.001D) break;
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }
                // Clearance changes page-top spacing and can change replayed content.
                // Recheck the complete group against any lower floating regions.
                measuredHeight = MeasureFlowBlocks(blocks);
                available = y - FootnoteBodyBottom;
            }
            if (!measuredHeight.HasValue &&
                (flow.Options.KeepTogether || flow.Options.OverflowBehavior != PdfFlowOverflowBehavior.Continue)) {
                throw new NotSupportedException("KeepTogether and non-continuing overflow behavior require flow content whose height can be determined before rendering. Remove the constraint or move dynamic, multi-column, deferred-table, canvas, or explicit page-boundary content outside the flow.");
            }

            bool cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + 0.001D;
            double? fullPageMeasuredHeight = measuredHeight.HasValue
                ? MeasureBlockSequenceAtFrameStart(blocks, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints())
                : null;
            RunningFrameAdmissionAttempt? runningAdmission = null;
            void RemeasureFlowFrame() {
                context = CreateFlowContext();
                if (flow.IsReplayable) blocks = MaterializeFlow(flow, context);
                measuredHeight = MeasureFlowBlocks(blocks);
                fullPageMeasuredHeight = measuredHeight.HasValue
                    ? MeasureBlockSequenceAtFrameStart(blocks, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints()) : null;
                available = y - FootnoteBodyBottom;
                beforeFloatClearanceY = y;
                cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + .001D;
            }
            void AdmitRunningFlowFrame() {
                while (flow.Options.KeepTogether && fullPageMeasuredHeight.HasValue &&
                    fullPageMeasuredHeight.Value > GetMaximumBlockContinuationHeight() + .001D &&
                    TryAdvanceForRunningContent(ref runningAdmission, fullPageMeasuredHeight.Value,
                        preserve => NewBlockFrame(preserveEmptyPage: preserve), contextualBody: flow.IsReplayable)) {
                    RemeasureFlowFrame();
                }
            }
            AdmitRunningFlowFrame();
            bool fitsFullPage = fullPageMeasuredHeight.HasValue &&
                                fullPageMeasuredHeight.Value <= GetMaximumBlockContinuationHeight() + 0.001D;
            // A fit against full-page capacity in an unused partial column does not finish the width search.
            bool widthAdaptiveKeep = flow.Options.KeepTogether && fullPageMeasuredHeight.HasValue && !fitsFullPage &&
                PdfFlowNestingRules.IsColumnFlowSupported(flow, allowTableOfContents: activeColumnFlow != null);
            bool needsWidthFit = widthAdaptiveKeep && CanFlowFitAnotherFlowWidth(blocks, width);
            bool moveForKeepTogether = flow.Options.KeepTogether && cannotFitCurrentPage && fitsFullPage;
            bool moveForOverflow = flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.MoveToNextPage && cannotFitCurrentPage && fitsFullPage;
            bool moveForMinimumHeight = flow.Options.MinimumRemainingHeight.ToPoints() > 0D &&
                available + 0.001D < flow.Options.MinimumRemainingHeight.ToPoints();
            while ((needsWidthFit || moveForKeepTogether || moveForOverflow || moveForMinimumHeight) &&
                   (needsWidthFit || ShouldAdvanceForBlockHeight(Math.Max(measuredHeight.GetValueOrDefault(), flow.Options.MinimumRemainingHeight.ToPoints())))) {
                NewBlockFrame();
                context = CreateFlowContext();
                if (flow.IsReplayable) {
                    blocks = MaterializeFlow(flow, context);
                }

                measuredHeight = MeasureFlowBlocks(blocks);
                fullPageMeasuredHeight = measuredHeight;
                // An ordinary keep/overflow move can enter a shorter running frame.
                AdmitRunningFlowFrame();
                available = y - FootnoteBodyBottom;
                beforeFloatClearanceY = y;
                cannotFitCurrentPage = measuredHeight.HasValue && measuredHeight.Value > available + 0.001D;
                fitsFullPage = measuredHeight.HasValue && measuredHeight.Value <= GetMaximumBlockContinuationHeight() + .001D;
                needsWidthFit = widthAdaptiveKeep && !fitsFullPage && CanFlowFitAnotherFlowWidth(blocks, width);
                moveForKeepTogether = flow.Options.KeepTogether && cannotFitCurrentPage && fitsFullPage;
                moveForOverflow = flow.Options.OverflowBehavior == PdfFlowOverflowBehavior.MoveToNextPage && cannotFitCurrentPage && fitsFullPage;
                moveForMinimumHeight = flow.Options.MinimumRemainingHeight.ToPoints() > 0D && available + .001D < flow.Options.MinimumRemainingHeight.ToPoints();
            }

            if (flow.Options.KeepTogether && fullPageMeasuredHeight.HasValue &&
                fullPageMeasuredHeight.Value > GetMaximumBlockContinuationHeight() + 0.001D) {
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

            if (flow.Options.KeepTogether && measuredHeight.HasValue)
                ReserveAtomicFootnoteGroup(blocks, () => MeasureFlowBlocks(blocks) ?? measuredHeight.Value,
                    () => { NewBlockFrame(); RemeasureFlowFrame(); AdmitRunningFlowFrame(); });
            int startPageNumber = pages.Count + 1;
            double startY = y;
            PdfCreateOptions startOptions = currentOpts;
            var paintedRegions = new FloatingFlowCapture();
            if (capture != null) activeFloatingFlowCaptures.Push(paintedRegions);
            try { ProcessBlocks(blocks, flow); }
            finally { if (capture != null) activeFloatingFlowCaptures.Pop(); }
            CaptureFlowRegions(capture, startPageNumber, startY, startOptions, paintedRegions);
        }

        private IReadOnlyList<IPdfBlock> MaterializeFlow(FlowBlock flow, PdfFlowContext context) {
            if (!flow.IsReplayable) return flow.Materialize(context);
            var key = new FlowMaterializationKey(flow, context);
            if (!deferredMaterializations.TryGetValue(key, out IReadOnlyList<IPdfBlock>? blocks)) {
                blocks = flow.Materialize(context);
                if (blocks.SelectMany(PdfFootnoteTraversal.References).Any())
                    throw new NotSupportedException("Page-associated footnote references are not supported in replayable flow. Use a static flow group.");
                deferredMaterializations.Add(key, blocks);
            }

            if (footnoteSourceOrdinals.Count > 0)
                foreach (IPdfBlock block in blocks) PdfDocument.ValidateFootnoteBodyProfile(block, currentOpts);
            return blocks;
        }

        private PdfFlowContext CreateFlowContext() {
            // Reserved notes reduce body room without advancing the frame cursor.
            return new PdfFlowContext(
                pages.Count + 1,
                y - FootnoteBodyBottom,
                GetFullPageContentHeight(),
                width,
                currentOpts.PageWidth.ToPoints(),
                currentOpts.PageHeight.ToPoints(),
                y >= GetCurrentFramePageStartY() - 0.001D);
        }

        private double GetFullPageContentHeight() {
            return GetCurrentFramePageStartY() - currentOpts.MarginBottom.ToPoints();
        }

        private double? MeasureFlowBlocks(IReadOnlyList<IPdfBlock> blocks) {
            return MeasureBlockSequence(blocks, currentOpts.MarginLeft.ToPoints(), width, currentOpts.DefaultFontSize.ToPoints());
        }

        private bool CanFlowFitAnotherFlowWidth(IReadOnlyList<IPdfBlock> blocks, double parentWidth) {
            return CanFitAnotherFixedFlowWidth(parentWidth, candidateWidth =>
                HasUsableContainerMeasurementWidths(blocks, candidateWidth, wholeContainer: true)
                    ? MeasureBlockSequenceAtFrameStart(blocks, currentOpts.MarginLeft.ToPoints(), candidateWidth, currentOpts.DefaultFontSize.ToPoints())
                    : null);
        }

        private void CaptureFlowRegions(PdfLayoutPositionCapture? capture, int startPageNumber, double startY, PdfCreateOptions startOptions, FloatingFlowCapture paintedRegions) {
            if (capture == null) {
                return;
            }

            void AddCapturedRegion(PdfLayoutRegion region) {
                var painted = paintedRegions.PaintedRegions.Where(item => item.PageNumber == region.PageNumber).ToList();
                if (painted.Count == 0) { capture.Add(region); return; }
                // A float-only group leaves its flow cursor unchanged. Do not let that
                // nominal frame extend the painted table's captured bounds, even when
                // pagination synthesized a completed-page frame.
                if (paintedRegions.FlowPages.Contains(region.PageNumber) && region.Height.ToPoints() > 0) painted.Add(region);
                double left = painted.Min(item => item.X.ToPoints()), bottom = painted.Min(item => item.Y.ToPoints());
                double right = painted.Max(item => item.X.ToPoints() + item.Width.ToPoints()), top = painted.Max(item => item.Y.ToPoints() + item.Height.ToPoints());
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
                AddCapturedRegion(new PdfLayoutRegion(startPageNumber, startOptions.MarginLeft.ToPoints(), bottom, startOptions.PageWidth.ToPoints() - startOptions.MarginLeft.ToPoints() - startOptions.MarginRight.ToPoints(), height));
                return;
            }

            AddCapturedRegion(new PdfLayoutRegion(
                startPageNumber,
                startOptions.MarginLeft.ToPoints(),
                startOptions.MarginBottom.ToPoints(),
                startOptions.PageWidth.ToPoints() - startOptions.MarginLeft.ToPoints() - startOptions.MarginRight.ToPoints(),
                Math.Max(0D, startY - startOptions.MarginBottom.ToPoints())));

            for (int pageNumber = startPageNumber + 1; pageNumber < endPageNumber; pageNumber++) {
                LayoutResult.Page completedPage = pages[pageNumber - 1];
                PdfCreateOptions options = completedPage.Options;
                AddCapturedRegion(new PdfLayoutRegion(
                    pageNumber,
                    options.MarginLeft.ToPoints(),
                    options.MarginBottom.ToPoints(),
                    options.PageWidth.ToPoints() - options.MarginLeft.ToPoints() - options.MarginRight.ToPoints(),
                    options.PageHeight.ToPoints() - options.MarginTop.ToPoints() - options.MarginBottom.ToPoints()));
            }

            if (currentPage != null) {
                AddCapturedRegion(new PdfLayoutRegion(
                    endPageNumber,
                    currentOpts.MarginLeft.ToPoints(),
                    y,
                    currentOpts.PageWidth.ToPoints() - currentOpts.MarginLeft.ToPoints() - currentOpts.MarginRight.ToPoints(),
                    Math.Max(0D, yStart - y)));
            }
        }
    }
}
