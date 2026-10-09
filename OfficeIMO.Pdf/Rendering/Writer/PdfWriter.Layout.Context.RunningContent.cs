using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly RunningContentImageAssets runningAssets;
        private readonly IReadOnlyList<PageNumberInfo>? previousRunningPages;
        private readonly int previousDocumentPages;
        private readonly Dictionary<(PdfOptions Source, double Top, double Bottom), PdfOptions> runningFrameOptions = new();

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
            double headerHeight = header == null ? 0D : MeasureRunningContent(source, header, headerBlocks!);
            double footerHeight = footer == null ? 0D : MeasureRunningContent(source, footer, footerBlocks!);
            double top = headerHeight <= 0D ? source.MarginTop
                : Math.Max(source.MarginTop, header!.DistanceFromEdge + headerHeight + header.BodyGap);
            double bottom = footerHeight <= 0D ? source.MarginBottom
                : Math.Max(source.MarginBottom, footer!.DistanceFromEdge + footerHeight + footer.BodyGap);
            if (source.PageHeight - top - bottom <= .001D)
                throw new InvalidOperationException("Running PDF content must leave a positive body content frame.");
            currentOpts = GetRunningFrameOptions(source, top, bottom);
            currentPage!.Options = currentOpts;
            currentPage.RunningContentContext = context;
            currentPage.HasDynamicRunningContent = header?.UsesPageContext == true || footer?.UsesPageContext == true;
            yStart = source.PageHeight - top;
            y = yStart;
            if (header != null && headerHeight > 0D)
                RenderRunningContent(source, headerBlocks!, source.PageHeight - header.DistanceFromEdge, headerHeight, isHeader: true);
            if (footer != null && footerHeight > 0D)
                RenderRunningContent(source, footerBlocks!, footer.DistanceFromEdge + footerHeight, footerHeight, isHeader: false);
        }

        private double MeasureRunningContent(PdfOptions source, PdfRunningContent definition, IReadOnlyList<IPdfBlock> blocks) {
            if (blocks.Count == 0) return 0D;
            double start = source.PageHeight - definition.DistanceFromEdge;
            if (start <= 0D) throw new InvalidOperationException("Running PDF content starts outside the page.");
            PdfOptions frame = GetRunningFrameOptions(source, definition.DistanceFromEdge, 0D);
            using var measure = new LayoutContext(frame, cancellationToken: cancellationToken,
                sharedPageContents: pageContents, isRunningContent: true);
            measure.currentPage = currentPage;
            measure.width = width;
            measure.yStart = start;
            measure.y = start;
            double? height = measure.MeasureBlockSequence(blocks, frame.MarginLeft, width, frame.DefaultFontSize);
            if (!height.HasValue)
                throw new NotSupportedException("Running PDF content must contain bounded, measurable flow; page breaks, deferred flow and page canvases are unsupported.");
            if (height.Value > start + .001D)
                throw new InvalidOperationException("Running PDF content exceeds the physical page height.");
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
            pageDirty = true;
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
