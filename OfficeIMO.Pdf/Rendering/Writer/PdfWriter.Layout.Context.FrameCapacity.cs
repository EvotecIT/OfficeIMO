namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private ColumnFlowScope? columnBalanceMeasurementScope;
        private PdfCreateOptions? measuredKeepPageOptions;
        private double? measuredKeepFrameTop;

        private (PdfCreateOptions Page, PdfCreateOptions Body, double Top, double Width)? MeasureKeepContinuationFrame() {
            if (pageFootnotesContinue) return null;
            var source = nextRunningPageOptions ?? currentPageBaseOptions;
            // The positioned content has already made this page part of the output.
            int nextPageIndex = pages.Count + 1;
            int previousVisible = currentVisiblePageNumber;
            if (currentPage?.ContinuedPageGroups is { } groups)
                foreach (var group in groups)
                    if (group.Options.HasExplicitPageNumberStart) previousVisible = group.Options.PageNumberStart;
            int nextVisible = ResolveNextVisiblePageNumber(nextPageIndex, firstPageOfGroup: false, previousVisible, source);
            PdfCreateOptions page = GetEffectivePageOptions(source, nextVisible);
            double frameX = page.MarginLeft.ToPoints();
            double frameWidth = page.PageWidth.ToPoints() - frameX - page.MarginRight.ToPoints();
            double topMargin = page.MarginTop.ToPoints(), bottom = page.MarginBottom.ToPoints();
            if (page.HasAnyRunningContent) {
                emittedPageGroupCounts.TryGetValue(currentPageGroupId, out int emittedGroupPages);
                int sectionPage = previousRunningPages != null && nextPageIndex < previousRunningPages.Count
                    ? previousRunningPages[nextPageIndex].VariantPageNumber : emittedGroupPages + 2;
                int variant = page.GetHeaderFooterVariantPageNumber(sectionPage, nextVisible);
                var header = page.GetRunningContentForPage(variant, isHeader: true);
                var footer = page.GetRunningContentForPage(variant, isHeader: false);
                // Contextual stories are materialized when their page actually starts.
                if (header?.UsesPageContext == true || footer?.UsesPageContext == true) return null;
                PdfRunningContentContext context = GetRunningPageContext();
                topMargin = MeasureStaticStoryExclusion(page, header, context, topMargin, frameWidth, isHeader: true);
                bottom = MeasureStaticStoryExclusion(page, footer, context, bottom, frameWidth, isHeader: false);
            }

            double top = page.PageHeight.ToPoints() - topMargin;
            for (int index = 0; index <= activeContainerScopes.Count; index++) {
                if (activeColumnFlow is { } columns && index == columns.ContainerDepth) frameWidth = columns.Widths[0];
                if (index == activeContainerScopes.Count) break;
                var scope = activeContainerScopes[index];
                var frame = ResolveContainerFrame(scope.Container, scope.Style, frameX, frameWidth);
                frameX = frame.X + scope.Style.PaddingX.ToPoints();
                frameWidth = frame.ContentWidth;
                bottom += scope.Style.FragmentBottomInset.ToPoints();
                top -= scope.Style.GetFragmentTopPadding(isContinuation: true);
            }

            // The body scope can carry its own fonts and styles inside a panel or column.
            PdfCreateOptions body = (activeContainerScopes.Count == 0 && activeColumnFlow == null ? page : currentOpts)
                .CreateRunningContentFrame(topMargin, bottom);
            body.MarginLeft = global::OfficeIMO.OfficeLength.Points(frameX);
            body.MarginRight = global::OfficeIMO.OfficeLength.Points(page.PageWidth.ToPoints() - frameX - frameWidth);
            return (page, body, top, frameWidth);
        }

        private double MeasureStaticStoryExclusion(PdfCreateOptions page, PdfRunningContent? story,
            PdfRunningContentContext context, double margin, double frameWidth, bool isHeader) {
            if (story == null) return margin;
            var blocks = MaterializeRunningContent(story, context);
            double? height = MeasureRunningContent(page, story, blocks, isHeader, out _, frameWidth);
            return height.HasValue ? GetRunningContentExclusion(story, height.Value, margin) : margin;
        }

        /// <summary>Uses physical continuation capacity while accounting for active or preflight-only container padding exactly once.</summary>
        private double GetMeasuredContinuationFrameHeight() {
            ColumnFlowScope? scope = columnBalanceMeasurementScope ?? activeColumnFlow;
            double height = scope == null ? yStart - currentOpts.MarginBottom.ToPoints() : scope.ParentYStart - scope.ParentOptions.MarginBottom.ToPoints();
            int activeDepth = columnBalanceMeasurementScope == null ? activeContainerScopes.Count : columnBalanceMeasurementScope.ContainerDepth;
            for (int index = 0; index < activeDepth; index++) {
                height -= activeContainerScopes[index].Style.GetFragmentTopPadding(isContinuation: true);
                if (scope != null && index >= scope.ContainerDepth) height -= activeContainerScopes[index].Style.FragmentBottomInset.ToPoints();
            }
            return Math.Max(0D, height - containerMeasurementTopPadding - containerMeasurementBottomInset);
        }

        /// <summary>Includes the full physical page available after columns that start partway down a page.</summary>
        private double GetMaximumTableContinuationFrameHeight(PdfTableStyle style, double currentFrameHeight) {
            if (style.Position != null || activeColumnFlow == null) return currentFrameHeight;
            return Math.Max(currentFrameHeight, GetMeasuredContinuationFrameHeight());
        }
    }
}
