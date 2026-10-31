using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly RunningContentImageAssets runningAssets;
        private readonly IReadOnlyList<PageNumberInfo>? previousRunningPages;
        private readonly int previousDocumentPages;
        private readonly Dictionary<(PdfCreateOptions Source, double Top, double Bottom), PdfCreateOptions> runningFrameOptions = new();
        private double? runningFirstBodyHeight;
        private bool runningFirstBodyAllowsFooterOverlap;
        private bool runningFirstBodyAllowsExactOverflow;

        // Imported exact-height body lines retain their authored baseline even
        // beyond the bottom margin. Admit only one line on an empty page or
        // column; decorated containers and running stories keep their bounds.
        private bool CanOverflowImportedExactLine(PdfParagraphStyle? style) =>
            !isRunningContent && activeContainerScopes.Count == 0 &&
            style?.LineSpacing?.FixedLineBoxBaselineOffset != null;

        private bool CanOverflowImportedExactFirstBlock(IPdfBlock block) {
            if (block is MultiColumnBlock columns) {
                foreach (IPdfBlock child in ExpandTransparentMeasurementBlocks(columns.Blocks)) {
                    if (IsNonVisualFlowMarker(child) || child is ColumnBreakBlock ||
                        child is SpacerBlock { Height: <= .001D }) continue;
                    return child is RichParagraphBlock paragraph && (CanOverflowImportedExactLine(EffectiveParagraphStyle(paragraph)) ||
                        paragraph.Runs.Any(run => run.InlineElement is PdfInlineImage { AllowFrameOverflow: true }));
                }
                return false;
            }
            return block is RichParagraphBlock exactParagraph && (CanOverflowImportedExactLine(EffectiveParagraphStyle(exactParagraph)) ||
                exactParagraph.Runs.Any(run => run.InlineElement is PdfInlineImage { AllowFrameOverflow: true }));
        }

        // A paragraph continuation can reclaim an optional imported header's
        // exclusion. Preserve footer and active-container reservations.
        private double GetMaximumParagraphContinuationHeight() {
            double height = GetMaximumBlockContinuationHeight();
            if (activeColumnFlow != null || currentPage?.RunningContentContext is not { } context) return height;
            int variant = currentOpts.GetHeaderFooterVariantPageNumber(context.SectionPageNumber, context.PageNumber);
            return currentOpts.GetRunningContentForPage(variant, isHeader: true)?.AllowTextFallback == true
                ? height + Math.Max(0D, currentOpts.MarginTop.ToPoints() - currentPageBaseOptions.MarginTop.ToPoints()) : height;
        }

        // Imported Word pages can place one body line beside a tall footer even
        // when its expanded exclusion leaves less than one line of body space.
        // Ordinary PDF flow and the physical body margins retain their bounds.
        private bool CanOverlapRunningFooterWithFirstLine(PdfParagraphStyle? style, double height) =>
            !isRunningContent && style?.AllowFirstLineOverlapWithRunningFooter == true && activeColumnFlow == null &&
            currentOpts.MarginBottom.ToPoints() > currentPageBaseOptions.MarginBottom.ToPoints() + .001D &&
            height <= GetCurrentFramePageStartY() - currentPageBaseOptions.MarginBottom.ToPoints() + .001D;

        private PdfCreateOptions GetRunningFrameOptions(PdfCreateOptions source, double top, double bottom) {
            var key = (source, top, bottom);
            if (!runningFrameOptions.TryGetValue(key, out PdfCreateOptions? frame)) {
                frame = source.CreateRunningContentFrame(top, bottom);
                runningFrameOptions.Add(key, frame);
            }
            return frame;
        }

        private PdfRunningContentContext GetRunningPageContext() {
            emittedPageGroupCounts.TryGetValue(currentPageGroupId, out int groupCount);
            int sectionPage = previousRunningPages != null && pages.Count < previousRunningPages.Count
                ? previousRunningPages[pages.Count].VariantPageNumber : groupCount + 1;
            int total = previousRunningPages != null && pages.Count < previousRunningPages.Count
                ? previousRunningPages[pages.Count].TotalPages : Math.Max(1, currentVisiblePageNumber);
            int sectionPages = previousRunningPages != null && pages.Count < previousRunningPages.Count
                ? previousRunningPages[pages.Count].SectionPages : sectionPage;
            return new PdfRunningContentContext(currentVisiblePageNumber, total, pages.Count + 1,
                Math.Max(1, previousDocumentPages), sectionPage, sectionPages, width, currentOpts.PageWidth.ToPoints(), currentOpts.PageHeight.ToPoints(),
                currentOpts.PageNumberStyle);
        }

        private IReadOnlyList<IPdfBlock> MaterializeRunningContent(PdfRunningContent content, PdfRunningContentContext context) {
            cancellationToken.ThrowIfCancellationRequested();
            // Static definitions already own their blocks. Dynamic stories are needed only
            // while measuring/rendering this page; retaining their graphs across pages and
            // stabilization passes multiplies document assets by the page count.
            return content.Materialize(context)
                ?? throw new InvalidOperationException("Running PDF content materialization returned null.");
        }

        private void PrepareAndRenderRunningContents() {
            PdfRunningContentContext context = GetRunningPageContext();
            int variant = currentOpts.GetHeaderFooterVariantPageNumber(context.SectionPageNumber, context.PageNumber);
            PdfRunningContent? header = currentOpts.GetRunningContentForPage(variant, isHeader: true);
            PdfRunningContent? footer = currentOpts.GetRunningContentForPage(variant, isHeader: false);
            if (header == null && footer == null) return;
            PdfCreateOptions source = currentOpts;
            IReadOnlyList<IPdfBlock>? headerBlocks = header == null ? null : MaterializeRunningContent(header, context);
            IReadOnlyList<IPdfBlock>? footerBlocks = footer == null ? null : MaterializeRunningContent(footer, context);
            bool usesPageContext = header?.UsesPageContext == true || footer?.UsesPageContext == true;
            bool headerHasNonFlowTablePaint = false, footerHasNonFlowTablePaint = false;
            double? measuredHeader = header == null ? 0D : MeasureRunningContent(source, header, headerBlocks!, true, out headerHasNonFlowTablePaint);
            double? measuredFooter = footer == null ? 0D : MeasureRunningContent(source, footer, footerBlocks!, false, out footerHasNonFlowTablePaint);
            bool fallbackHeader = !measuredHeader.HasValue;
            bool fallbackFooter = !measuredFooter.HasValue;
            if (fallbackHeader) header = null;
            if (fallbackFooter) footer = null;
            double headerHeight = measuredHeader ?? 0D;
            double footerHeight = measuredFooter ?? 0D;
            double top = Exclusion(header, headerHeight, source.MarginTop.ToPoints());
            double bottom = Exclusion(footer, footerHeight, source.MarginBottom.ToPoints());
            if (header?.AllowTextFallback == true && top > source.MarginTop.ToPoints() + .001D) {
                double? footerOverlapHeight;
                double firstBodyHeight;
                bool allowsExactOverflow;
                if (runningFirstBodyHeight is double continuationHeight) {
                    // Paragraph flow already measured this actual continuation
                    // line, including its active-container reservation.
                    firstBodyHeight = continuationHeight;
                    footerOverlapHeight = runningFirstBodyAllowsFooterOverlap ? continuationHeight : null;
                    allowsExactOverflow = runningFirstBodyAllowsExactOverflow;
                } else firstBodyHeight = MeasureRunningContentFirstBodyHeight(source, out footerOverlapHeight, out allowsExactOverflow);
                // A supported imported paragraph can retain its header and put
                // its first line beside the footer. Admission must use the same
                // physical bottom bound as paragraph flow before dropping it.
                bool fitsBesideFooter = footerOverlapHeight.HasValue && bottom > source.MarginBottom.ToPoints() + .001D &&
                    footerOverlapHeight.Value <= source.PageHeight.ToPoints() - top - source.MarginBottom.ToPoints() + .001D;
                if (!allowsExactOverflow && !fitsBesideFooter && firstBodyHeight <= source.PageHeight.ToPoints() - source.MarginTop.ToPoints() - source.MarginBottom.ToPoints() + .001D &&
                    source.PageHeight.ToPoints() - top - bottom + .001D < firstBodyHeight) {
                    fallbackHeader = true;
                    header = null;
                    headerHeight = 0D;
                    top = source.MarginTop.ToPoints();
                }
            }
            if (source.PageHeight.ToPoints() - top - bottom <= .001D) {
                // Ordinary imported stories retain styled text projections. Use
                // them when this page's actual fields and opposite story leave
                // no body frame; required native flow keeps its strict contract.
                if (header?.AllowTextFallback == true) {
                    fallbackHeader = true;
                    header = null;
                    headerHeight = 0D;
                    top = source.MarginTop.ToPoints();
                }
                if (footer?.AllowTextFallback == true) {
                    fallbackFooter = true;
                    footer = null;
                    footerHeight = 0D;
                    bottom = source.MarginBottom.ToPoints();
                }
            }
            if (source.PageHeight.ToPoints() - top - bottom <= .001D)
                throw new InvalidOperationException("Running PDF content must leave a positive body content frame.");
            // A fallback is local to this page and variant. Never remove a
            // definition from the shared source or cached geometry frames.
            bool retainHeaderProjection = header?.RetainTextProjection == true;
            bool retainFooterProjection = footer?.RetainTextProjection == true;
            currentOpts = fallbackHeader || fallbackFooter || retainHeaderProjection || retainFooterProjection
                ? source.CreateRunningContentFrame(top, bottom) : GetRunningFrameOptions(source, top, bottom);
            if (fallbackHeader) currentOpts.ClearRunningContentForPage(variant, isHeader: true);
            if (fallbackFooter) currentOpts.ClearRunningContentForPage(variant, isHeader: false);
            // Geometry and visible text have different owners here. Keep the
            // measured body frame, but draw projected text from its original
            // physical margin, without mutating shared or cached options.
            if (retainHeaderProjection) {
                currentOpts.HeaderOffsetY += global::OfficeIMO.OfficeLength.Points(top - source.MarginTop.ToPoints());
            }
            if (retainFooterProjection) {
                currentOpts.FooterOffsetY += global::OfficeIMO.OfficeLength.Points(bottom - source.MarginBottom.ToPoints());
            }
            currentPage!.Options = currentOpts;
            currentPage.RunningContentContext = context;
            currentPage.HasDynamicRunningContent = usesPageContext;
            yStart = source.PageHeight.ToPoints() - top;
            y = yStart;
            if (header != null && (headerHeight > 0D || headerHasNonFlowTablePaint))
                RenderRunningContent(source, headerBlocks!, source.PageHeight.ToPoints() - header.DistanceFromEdge, headerHeight,
                    isHeader: true, headerHasNonFlowTablePaint);
            if (footer != null && (footerHeight > 0D || footerHasNonFlowTablePaint))
                RenderRunningContent(source, footerBlocks!, footer.DistanceFromEdge + footerHeight, footerHeight,
                    isHeader: false, footerHasNonFlowTablePaint);

            static double Exclusion(PdfRunningContent? content, double height, double physicalMargin) =>
                content == null || (height <= 0D && !content.ReserveDistanceWhenEmpty) ? physicalMargin
                    : Math.Max(physicalMargin, content.DistanceFromEdge + height + content.BodyGap);
        }

        private double MeasureRunningContentFirstBodyHeight(PdfCreateOptions source, out double? footerOverlapHeight, out bool allowsExactOverflow) {
            double? firstVisualFooterOverlapHeight = null;
            bool firstVisualAllowsExactOverflow = false;
            footerOverlapHeight = null;
            allowsExactOverflow = false;
            for (int depth = activeBlockSequences.Count - 1; depth >= 0; depth--) {
                BlockSequenceScope sequence = activeBlockSequences[depth];
                int start = sequence.Index + (depth == activeBlockSequences.Count - 1 ? 0 : 1);
                if (depth == activeBlockSequences.Count - 1 && start < sequence.Blocks.Count &&
                    sequence.Blocks[start] is PageBreakBlock) start++;
                var frame = GetRunningBodyContinuationFrame(source, depth);
                double? height = FirstVisual(PendingBlocks(sequence.Blocks, start), frame.X, frame.Width,
                    GetClosingContainerPadding(depth));
                if (height.HasValue) {
                    allowsExactOverflow = firstVisualAllowsExactOverflow;
                    footerOverlapHeight = firstVisualFooterOverlapHeight.HasValue
                        ? firstVisualFooterOverlapHeight.Value + frame.ReservedHeight : null;
                    return height.Value <= .001D ? height.Value : height.Value + frame.ReservedHeight;
                }
                // Do not look through the end of an explicit page/section.
                if (depth > 0 && activeBlockSequences[depth - 1].Blocks[activeBlockSequences[depth - 1].Index] is PageBlock) break;
            }
            return 0D;

            double? FirstVisual(IEnumerable<IPdfBlock> blocks, double frameX, double frameWidth, double closingPadding) {
                using var pending = ExpandTransparentMeasurementBlocks(blocks).GetEnumerator();
                while (pending.MoveNext()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    IPdfBlock block = pending.Current;
                    if (block is BookmarkBlock or RunningContentSectionBlock or ColumnBreakBlock) continue;
                    if (block is PageBreakBlock) return 0D;
                    if (block is PageBlock page) return FirstVisual(page.Blocks, frameX, frameWidth, closingPadding) ?? 0D;
                    double height = MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, source.DefaultFontSize.ToPoints(),
                        allowTableFragments: true, suppressParagraphSpacingBefore: true);
                    if (height <= .001D && block is ContainerBlock or SpacerBlock) continue;
                    double paragraphClosingPadding = 0D;
                    if (closingPadding > 0D) {
                        bool closesSequence = true;
                        while (pending.MoveNext()) {
                            cancellationToken.ThrowIfCancellationRequested();
                            if (!IsNonVisualFlowMarker(pending.Current)) { closesSequence = false; break; }
                        }
                        if (closesSequence) {
                            paragraphClosingPadding = closingPadding;
                            height = MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, source.DefaultFontSize.ToPoints(),
                                allowTableFragments: true, suppressParagraphSpacingBefore: true, closingPadding: closingPadding);
                        }
                    }
                    if (activeColumnFlow == null && block is RichParagraphBlock paragraph &&
                        EffectiveParagraphStyle(paragraph)?.AllowFirstLineOverlapWithRunningFooter == true)
                        firstVisualFooterOverlapHeight = MeasureParagraphFirstVisual(paragraph, frameX, frameWidth,
                            source.DefaultFontSize.ToPoints(), suppressSpacingBefore: true, closingPadding: paragraphClosingPadding,
                            reserveSingleLine: true).Height;
                    firstVisualAllowsExactOverflow = CanOverflowImportedExactFirstBlock(block);
                    return height;
                }
                return null;
            }

            static IEnumerable<IPdfBlock> PendingBlocks(IList<IPdfBlock> blocks, int start) {
                for (int index = start; index < blocks.Count; index++) yield return blocks[index];
            }
        }

        private double? MeasureRunningContent(PdfCreateOptions source, PdfRunningContent definition, IReadOnlyList<IPdfBlock> blocks,
            bool isHeader, out bool hasNonFlowTablePaint) {
            hasNonFlowTablePaint = false;
            if (blocks.Count == 0) return 0D;
            double start = source.PageHeight.ToPoints() - definition.DistanceFromEdge;
            if (start <= 0D) {
                if (definition.AllowTextFallback) return null;
                throw new InvalidOperationException("Running PDF content starts outside the page.");
            }
            PdfCreateOptions frame = GetRunningFrameOptions(source, definition.DistanceFromEdge, 0D);
            using var measure = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: true,
                drawingPatternAnnotations: drawingPatternAnnotations, pictureAnchors: pictureAnchors);
            measure.currentPage = currentPage;
            measure.width = width;
            measure.yStart = start;
            measure.y = start;
            measure.SetPictureStory(isHeader ? "Header" : "Footer", PicturePage, start);
            double? height = measure.MeasureBlockSequence(blocks, frame.MarginLeft.ToPoints(), width, frame.DefaultFontSize.ToPoints());
            hasNonFlowTablePaint = measure.measuredNonFlowTablePaint;
            if (!height.HasValue) {
                if (definition.AllowTextFallback) return null;
                throw new NotSupportedException("Running PDF content must contain bounded, measurable flow; page breaks, deferred flow and page canvases are unsupported.");
            }
            double paintHeight = Math.Max(height.Value, start - measure.measuredRunningPaintBottom);
            // A footer grows upward from its edge distance. Reserve only the
            // extra origin height needed to keep restored overlay paint on-page.
            double reservedHeight = isHeader ? height.Value
                : Math.Max(height.Value, paintHeight - definition.DistanceFromEdge);
            if (paintHeight > start + .001D || reservedHeight > start + .001D) {
                if (definition.AllowTextFallback) return null;
                throw new InvalidOperationException("Running PDF content exceeds the physical page height.");
            }
            return reservedHeight;
        }

        private void RenderRunningContent(PdfCreateOptions source, IReadOnlyList<IPdfBlock> blocks, double start, double height,
            bool isHeader, bool hasNonFlowTablePaint) {
            // Match the existing continuation tolerance: exact sums can differ by a
            // floating-point ulp when the final paragraph is tested for available room.
            // Overlay tables paint below the unchanged flow cursor. Let that
            // paint use the physical page while the final y check still enforces
            // the measured flow reservation. Ordinary stories keep their frame.
            PdfCreateOptions frame = GetRunningFrameOptions(source, source.PageHeight.ToPoints() - start,
                hasNonFlowTablePaint ? 0D : Math.Max(0D, start - height - .001D));
            using var story = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: true,
                drawingPatternAnnotations: drawingPatternAnnotations, pictureAnchors: pictureAnchors);
            story.currentPage = currentPage;
            story.width = width;
            story.yStart = start;
            story.y = start;
            story.SetPictureStory(isHeader ? "Header" : "Footer", PicturePage, start);
            story.currentVisiblePageNumber = currentVisiblePageNumber;
            story.currentPageGroupId = currentPageGroupId;
            int firstImage = currentPage!.Images.Count;
            story.ProcessBlocks(blocks);
            for (int index = firstImage; index < currentPage.Images.Count; index++)
                runningAssets.Share(currentPage.Images[index], cancellationToken);
            if (story.y < start - height - .001D)
                throw new InvalidOperationException("Running PDF content exceeded its measured frame.");
            if (story.behindTextCanvases.Count > 0) {
                var behind = new StringBuilder();
                foreach (var canvas in story.behindTextCanvases.OrderBy(item => item.ZOrder)) behind.Append(canvas.Content);
                story.sb.Insert(0, behind.ToString());
            }
            usedBold |= story.usedBold;
            usedItalic |= story.usedItalic;
            usedBoldItalic |= story.usedBoldItalic;
            sb.Append("q\n0 0 ").Append(source.PageWidth.ToPoints().ToString("0.###", CultureInfo.InvariantCulture)).Append(' ')
                .Append(source.PageHeight.ToPoints().ToString("0.###", CultureInfo.InvariantCulture)).Append(" re W n\n")
                .Append("/Artifact << /Type /Pagination /Subtype /").Append(isHeader ? "Header" : "Footer")
                .Append(" >> BDC\n").Append(story.sb).Append("EMC\nQ\n");
            // Running stories make a page printable but do not belong to its
            // body section. The first body content owns shared-page fields.
            pageHasContent = true;
        }
    }

    private static bool RunningContentContextsMatch(LayoutResult result, IReadOnlyList<PageNumberInfo> infos) {
        for (int index = 0; index < result.Pages.Count; index++) {
            LayoutResult.Page page = result.Pages[index];
            if (!page.HasDynamicRunningContent) continue;
            PdfRunningContentContext context = page.RunningContentContext!;
            PageNumberInfo info = infos[index];
            if (context.PageNumber != info.PageNumber || context.TotalPages != info.TotalPages ||
                context.DocumentPages != result.Pages.Count || context.DocumentPageNumber != index + 1 ||
                context.SectionPageNumber != info.VariantPageNumber || context.SectionPages != info.SectionPages) return false;
        }
        return true;
    }
}
