using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly RunningContentImageAssets runningAssets;
        private readonly IReadOnlyList<PageNumberInfo>? previousRunningPages;
        private readonly int previousDocumentPages;
        private readonly Dictionary<(PdfOptions Source, double Top, double Bottom), PdfOptions> runningFrameOptions = new();
        private double? runningFirstBodyHeight;
        private bool runningFirstBodyAllowsFooterOverlap;

        // A paragraph continuation can reclaim an optional imported header's
        // exclusion. Preserve footer and active-container reservations.
        private double GetMaximumParagraphContinuationHeight() {
            double height = GetMaximumBlockContinuationHeight();
            if (activeColumnFlow != null || currentPage?.RunningContentContext is not { } context) return height;
            int variant = currentOpts.GetHeaderFooterVariantPageNumber(context.SectionPageNumber, context.PageNumber);
            return currentOpts.GetRunningContentForPage(variant, isHeader: true)?.AllowTextFallback == true
                ? height + Math.Max(0D, currentOpts.MarginTop - currentPageBaseOptions.MarginTop) : height;
        }

        // Imported Word pages can place one body line beside a tall footer even
        // when its expanded exclusion leaves less than one line of body space.
        // Ordinary PDF flow and the physical body margins retain their bounds.
        private bool CanOverlapRunningFooterWithFirstLine(PdfParagraphStyle? style, double height) =>
            !isRunningContent && style?.AllowFirstLineOverlapWithRunningFooter == true && activeColumnFlow == null &&
            currentOpts.MarginBottom > currentPageBaseOptions.MarginBottom + .001D &&
            height <= GetCurrentFramePageStartY() - currentPageBaseOptions.MarginBottom + .001D;

        private PdfOptions GetRunningFrameOptions(PdfOptions source, double top, double bottom) {
            var key = (source, top, bottom);
            if (!runningFrameOptions.TryGetValue(key, out PdfOptions? frame)) {
                frame = source.CreateRunningContentFrame(top, bottom);
                runningFrameOptions.Add(key, frame);
            }
            return frame;
        }

        private PdfRunningContentContext GetRunningPageContext() {
            emittedPageGroupCounts.TryGetValue(currentPageGroupId, out int groupCount);
            int sectionPage = groupCount + 1;
            int total = previousRunningPages != null && pages.Count < previousRunningPages.Count
                ? previousRunningPages[pages.Count].TotalPages : Math.Max(1, currentVisiblePageNumber);
            int sectionPages = previousRunningPages != null && pages.Count < previousRunningPages.Count
                ? previousRunningPages[pages.Count].SectionPages : sectionPage;
            return new PdfRunningContentContext(currentVisiblePageNumber, total, pages.Count + 1,
                Math.Max(1, previousDocumentPages), sectionPage, sectionPages, width, currentOpts.PageWidth, currentOpts.PageHeight,
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
            PdfOptions source = currentOpts;
            IReadOnlyList<IPdfBlock>? headerBlocks = header == null ? null : MaterializeRunningContent(header, context);
            IReadOnlyList<IPdfBlock>? footerBlocks = footer == null ? null : MaterializeRunningContent(footer, context);
            bool usesPageContext = header?.UsesPageContext == true || footer?.UsesPageContext == true;
            double? measuredHeader = header == null ? 0D : MeasureRunningContent(source, header, headerBlocks!);
            double? measuredFooter = footer == null ? 0D : MeasureRunningContent(source, footer, footerBlocks!);
            bool fallbackHeader = !measuredHeader.HasValue;
            bool fallbackFooter = !measuredFooter.HasValue;
            if (fallbackHeader) header = null;
            if (fallbackFooter) footer = null;
            double headerHeight = measuredHeader ?? 0D;
            double footerHeight = measuredFooter ?? 0D;
            double top = Exclusion(header, headerHeight, source.MarginTop);
            double bottom = Exclusion(footer, footerHeight, source.MarginBottom);
            if (header?.AllowTextFallback == true && top > source.MarginTop + .001D) {
                double? footerOverlapHeight;
                double firstBodyHeight;
                if (runningFirstBodyHeight is double continuationHeight) {
                    // Paragraph flow already wrapped and measured this actual
                    // continuation line, including its container reservation.
                    firstBodyHeight = continuationHeight;
                    footerOverlapHeight = runningFirstBodyAllowsFooterOverlap ? continuationHeight : null;
                } else firstBodyHeight = MeasureRunningContentFirstBodyHeight(source, out footerOverlapHeight);
                // A supported imported paragraph can retain its header and put
                // its first line beside the footer. Admission must use the same
                // physical bottom bound as paragraph flow before dropping it.
                bool fitsBesideFooter = footerOverlapHeight.HasValue && bottom > source.MarginBottom + .001D &&
                    footerOverlapHeight.Value <= source.PageHeight - top - source.MarginBottom + .001D;
                if (!fitsBesideFooter && firstBodyHeight <= source.PageHeight - source.MarginTop - source.MarginBottom + .001D &&
                    source.PageHeight - top - bottom + .001D < firstBodyHeight) {
                    fallbackHeader = true;
                    header = null;
                    headerHeight = 0D;
                    top = source.MarginTop;
                }
            }
            if (source.PageHeight - top - bottom <= .001D) {
                // Ordinary imported stories retain styled text projections. Use
                // them when this page's actual fields and opposite story leave
                // no body frame; required native flow keeps its strict contract.
                if (header?.AllowTextFallback == true) {
                    fallbackHeader = true;
                    header = null;
                    headerHeight = 0D;
                    top = source.MarginTop;
                }
                if (footer?.AllowTextFallback == true) {
                    fallbackFooter = true;
                    footer = null;
                    footerHeight = 0D;
                    bottom = source.MarginBottom;
                }
            }
            if (source.PageHeight - top - bottom <= .001D)
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
                currentOpts.ClearRunningContentForPage(variant, isHeader: true);
                currentOpts.HeaderOffsetY += top - source.MarginTop;
            }
            if (retainFooterProjection) {
                currentOpts.ClearRunningContentForPage(variant, isHeader: false);
                currentOpts.FooterOffsetY += bottom - source.MarginBottom;
            }
            currentPage!.Options = currentOpts;
            currentPage.RunningContentContext = context;
            currentPage.HasDynamicRunningContent = usesPageContext;
            yStart = source.PageHeight - top;
            y = yStart;
            if (header != null && headerHeight > 0D)
                RenderRunningContent(source, headerBlocks!, source.PageHeight - header.DistanceFromEdge, headerHeight, isHeader: true);
            if (footer != null && footerHeight > 0D)
                RenderRunningContent(source, footerBlocks!, footer.DistanceFromEdge + footerHeight, footerHeight, isHeader: false);

            static double Exclusion(PdfRunningContent? content, double height, double physicalMargin) =>
                content == null || (height <= 0D && !content.ReserveDistanceWhenEmpty) ? physicalMargin
                    : Math.Max(physicalMargin, content.DistanceFromEdge + height + content.BodyGap);
        }

        private double MeasureRunningContentFirstBodyHeight(PdfOptions source, out double? footerOverlapHeight) {
            double? firstVisualFooterOverlapHeight = null;
            footerOverlapHeight = null;
            for (int depth = activeBlockSequences.Count - 1; depth >= 0; depth--) {
                BlockSequenceScope sequence = activeBlockSequences[depth];
                int start = sequence.Index + (depth == activeBlockSequences.Count - 1 ? 0 : 1);
                if (depth == activeBlockSequences.Count - 1 && start < sequence.Blocks.Count &&
                    sequence.Blocks[start] is PageBreakBlock) start++;
                var frame = GetRunningBodyContinuationFrame(source, depth);
                double? height = FirstVisual(PendingBlocks(sequence.Blocks, start), frame.X, frame.Width,
                    GetClosingContainerPadding(depth));
                if (height.HasValue) {
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
                    double height = MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, source.DefaultFontSize,
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
                            height = MeasureNextBlockFirstVisualHeight(block, frameX, frameWidth, source.DefaultFontSize,
                                allowTableFragments: true, suppressParagraphSpacingBefore: true, closingPadding: closingPadding);
                        }
                    }
                    if (activeColumnFlow == null && block is RichParagraphBlock paragraph &&
                        EffectiveParagraphStyle(paragraph)?.AllowFirstLineOverlapWithRunningFooter == true)
                        firstVisualFooterOverlapHeight = MeasureParagraphFirstVisual(paragraph, frameX, frameWidth,
                            source.DefaultFontSize, suppressSpacingBefore: true, closingPadding: paragraphClosingPadding,
                            reserveSingleLine: true).Height;
                    return height;
                }
                return null;
            }

            static IEnumerable<IPdfBlock> PendingBlocks(IList<IPdfBlock> blocks, int start) {
                for (int index = start; index < blocks.Count; index++) yield return blocks[index];
            }
        }

        private double? MeasureRunningContent(PdfOptions source, PdfRunningContent definition, IReadOnlyList<IPdfBlock> blocks) {
            if (blocks.Count == 0) return 0D;
            double start = source.PageHeight - definition.DistanceFromEdge;
            if (start <= 0D) {
                if (definition.AllowTextFallback) return null;
                throw new InvalidOperationException("Running PDF content starts outside the page.");
            }
            PdfOptions frame = GetRunningFrameOptions(source, definition.DistanceFromEdge, 0D);
            using var measure = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: true);
            measure.currentPage = currentPage;
            measure.width = width;
            measure.yStart = start;
            measure.y = start;
            double? height = measure.MeasureBlockSequence(blocks, frame.MarginLeft, width, frame.DefaultFontSize);
            if (!height.HasValue) {
                if (definition.AllowTextFallback) return null;
                throw new NotSupportedException("Running PDF content must contain bounded, measurable flow; page breaks, deferred flow and page canvases are unsupported.");
            }
            if (height.Value > start + .001D) {
                if (definition.AllowTextFallback) return null;
                throw new InvalidOperationException("Running PDF content exceeds the physical page height.");
            }
            return height.Value;
        }

        private void RenderRunningContent(PdfOptions source, IReadOnlyList<IPdfBlock> blocks, double start, double height, bool isHeader) {
            // Match the existing continuation tolerance: exact sums can differ by a
            // floating-point ulp when the final paragraph is tested for available room.
            PdfOptions frame = GetRunningFrameOptions(source, source.PageHeight - start, Math.Max(0D, start - height - .001D));
            using var story = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: true);
            story.currentPage = currentPage;
            story.width = width;
            story.yStart = start;
            story.y = start;
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
            sb.Append("q\n0 0 ").Append(source.PageWidth.ToString("0.###", CultureInfo.InvariantCulture)).Append(' ')
                .Append(source.PageHeight.ToString("0.###", CultureInfo.InvariantCulture)).Append(" re W n\n")
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
