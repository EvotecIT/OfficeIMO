namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private HtmlRenderDocument RenderPaged(IReadOnlyList<HtmlRenderFlowBlock> blocks) {
        _runningStringValues.Clear();
        _currentPageRunningStringAssignments.Clear();
        _currentRunningStringPage = new HtmlCssRunningStringPageContext(_runningStringValues);
        string? currentPageName = blocks.Count > 0 ? blocks[0].PageName : null;
        HtmlCssPageGeometry pageGeometry = _pageRules.ResolveGeometry(1, currentPageName, _options);
        SetActivePageGeometry(pageGeometry);
        double pageWidth = pageGeometry.Width;
        double pageHeight = pageGeometry.Height;
        double contentHeight = ResolvePageBodyContentHeight(1, pageGeometry);
        ValidateSurface(pageWidth, pageHeight);
        PrepareGlobalPositionedRequests(
            includeRoot: true,
            pageWidth,
            pageHeight,
            pageGeometry.ContentWidth,
            Math.Max(1D, contentHeight));
        BuildRootStackingPaintOrders(blocks);

        var pages = new List<HtmlRenderPage>();
        var visuals = CreatePageVisuals(pageWidth, pageHeight, pageGeometry);
        double y = ResolvePageBodyTop(1, pageGeometry);
        void BeginPage(string? pageName) {
            pageGeometry = _pageRules.ResolveGeometry(pages.Count + 1, pageName, _options);
            SetActivePageGeometry(pageGeometry);
            pageWidth = pageGeometry.Width;
            pageHeight = pageGeometry.Height;
            contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
            ValidateSurface(pageWidth, pageHeight);
            visuals = CreatePageVisuals(pageWidth, pageHeight, pageGeometry);
            y = ResolvePageBodyTop(pages.Count + 1, pageGeometry);
        }
        for (int index = 0; index < blocks.Count; index++) {
            CheckCancellation();
            HtmlRenderFlowBlock block = blocks[index];
            bool hasPageContent = y > ResolvePageBodyTop(pages.Count + 1, pageGeometry) + 0.0001D;
            if (!string.Equals(currentPageName, block.PageName, StringComparison.Ordinal)) {
                if (hasPageContent) CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(block.PageName);
                block = RelayoutTopLevelBlockForPage(block, pageGeometry).ForPagination();
                hasPageContent = false;
            } else {
                block = RelayoutTopLevelBlockForPage(block, pageGeometry).ForPagination();
            }

            if (!hasPageContent) currentPageName = block.PageName;
            HtmlPageBreakTarget breakBefore = ResolveForcedBreakAt(block.ForcedBreaks, 0D);
            if (breakBefore == HtmlPageBreakTarget.None && IsForcedPageBreak(block.BreakBefore)) breakBefore = block.BreakBefore;
            if (breakBefore != HtmlPageBreakTarget.None) {
                ApplyBreakBefore(breakBefore, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                pageWidth = pageGeometry.Width;
                pageHeight = pageGeometry.Height;
                contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                hasPageContent = y > ResolvePageBodyTop(pages.Count + 1, pageGeometry) + 0.0001D;
                currentPageName = block.PageName;
                block = RelayoutTopLevelBlockForPage(block, pageGeometry).ForPagination();
            }

            double remainingHeight = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
            HtmlRenderFlowBlock floatAwareBlock = block;
            bool deferredFloat = false;
            if (block.Height > remainingHeight + 0.0001D && remainingHeight > 0.0001D && !HasInternalForcedBreak(block)) {
                HtmlCssPageGeometry nextPageGeometry = _pageRules.ResolveGeometry(pages.Count + 2, block.PageName, _options);
                if (SamePageGeometry(pageGeometry, nextPageGeometry)
                    && Math.Abs(contentHeight - ResolvePageBodyContentHeight(pages.Count + 2, nextPageGeometry)) <= 0.0001D) {
                    deferredFloat = TryRelayoutBlockForDeferredFloat(
                        block, pageGeometry, remainingHeight, contentHeight, out floatAwareBlock);
                }
            }
            if (deferredFloat) block = floatAwareBlock.ForPagination();

            if (hasPageContent
                && !deferredFloat
                && !HasInternalForcedBreak(block)
                && ShouldMoveTopLevelKeepWithNext(blocks, index, block, remainingHeight, pageGeometry, pages.Count + 2)) {
                CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(block.PageName);
                currentPageName = block.PageName;
                block = RelayoutTopLevelBlockForPage(block, pageGeometry).ForPagination();
                hasPageContent = false;
            }

            if (block.Height <= contentHeight
                && hasPageContent
                && !deferredFloat
                && !HasInternalForcedBreak(block)
                && y + block.Height > ResolvePageBodyBottom(pages.Count + 1, pageGeometry)
                && (block.AvoidBreakInside
                    || FindFragmentEnd(block, 0D, remainingHeight, fullPageHeight: contentHeight) <= 0.0001D)) {
                CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(block.PageName);
                currentPageName = block.PageName;
                block = RelayoutTopLevelBlockForPage(block, pageGeometry).ForPagination();
            }

            if (TryRelayoutConstrainedOverflow(block,
                ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y, contentHeight,
                pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock constrainedBlock))
                block = constrainedBlock.ForPagination();

            // Auto-height columns need the physical capacity even when their
            // initial balanced source box fits. Run after whole-block keep
            // decisions so the mapping starts on the page that owns the box.
            bool columnAware = false;
            while (TryRelayoutBlockForColumns(block, 0D,
                ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y, contentHeight,
                pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock columnBlock)) {
                block = columnBlock.ForPagination();
                columnAware = true;
            }
            if (columnAware && hasPageContent) {
                HtmlPageBreakTarget columnBreak = ResolveForcedBreakAt(block.ForcedBreaks, 0D);
                if (columnBreak != HtmlPageBreakTarget.None) {
                    ApplyBreakBefore(columnBreak, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                    pageWidth = pageGeometry.Width;
                    pageHeight = pageGeometry.Height;
                    contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                }
            }

            if (block.Height <= ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y && !HasInternalForcedBreak(block)) {
                AddTranslatedVisuals(visuals, block.Visuals, pageGeometry.Margins.Left, y, block);
                RecordRunningStringAssignments(block, 0D, block.Height, y);
                y += block.Height;
            } else {
                block = QualifySharedTableRepetition(block);
                double blockOffset = 0D;
                while (blockOffset < block.Height - 0.0001D) {
                    CheckCancellation();
                    double flexAvailable = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
                    if (TryRelayoutBlockForColumns(block, blockOffset, flexAvailable, contentHeight,
                        pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock columnBlock)) {
                        block = QualifySharedTableRepetition(columnBlock.ForPagination());
                        continue;
                    }
                    if (TryRelayoutBlockForParallelTables(block, blockOffset, flexAvailable, contentHeight,
                        pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock parallelBlock)) {
                        block = QualifySharedTableRepetition(parallelBlock.ForPagination());
                        continue;
                    }
                    if (!_pagedFlexAlignedBlocks.Contains(block)
                        && flexAvailable > 0.0001D
                        && block.Height > blockOffset + flexAvailable + 0.0001D
                        && TryRelayoutBlockForFlexPagination(block, blockOffset, flexAvailable, contentHeight,
                            pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock alignedFlexBlock)) {
                        block = QualifySharedTableRepetition(alignedFlexBlock.ForPagination());
                        _pagedFlexAlignedBlocks.Add(block);
                    }
                    double rawAvailable = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
                    PagedBlockFragment? plan = PlanPagedFragment(block, blockOffset, rawAvailable, contentHeight,
                        canForceFragment: y <= ResolvePageBodyTop(pages.Count + 1, pageGeometry) + 0.0001D
                            && _pageFloatPlan.ReservedHeight(pages.Count + 1, "top") <= 0D
                            && _pageFloatPlan.ReservedHeight(pages.Count + 1, "bottom") <= 0D);
                    if (plan == null) {
                        CommitPage(pages, visuals, pageGeometry, currentPageName);
                        BeginPage(currentPageName);
                        continue;
                    }
                    ReportPagedFragmentLoss(block, plan);
                    HtmlRenderContinuationGroup? continuationGroup = plan.Header;
                    HtmlRenderTrailingGroup? trailingGroup = plan.Footer;
                    bool repeatContinuation = continuationGroup != null;
                    bool repeatTrailing = trailingGroup != null;
                    double continuationHeight = plan.HeaderHeight;
                    double trailingHeight = plan.FooterHeight;
                    double fragmentEnd = plan.End;
                    var fragmentVisuals = new List<HtmlRenderVisual>();

                    ReportSharedTableRepetitionLoss(block, fragmentEnd);
                    if (repeatContinuation) {
                        foreach (HtmlRenderVisual visual in continuationGroup!.Visuals) {
                            fragmentVisuals.Add(visual.Translate(0D, y, fragmentVisuals.Count));
                        }
                        y += continuationHeight;
                    }

                    IReadOnlyList<HtmlRenderVisual> fragment = SliceBlockVisuals(block, blockOffset, fragmentEnd,
                        ResolvePagedBackgroundWindow(blockOffset, plan));
                    foreach (HtmlRenderVisual visual in fragment) {
                        fragmentVisuals.Add(visual.Translate(0D, y, fragmentVisuals.Count));
                    }
                    RecordRunningStringAssignments(block, blockOffset, fragmentEnd, y);
                    y += fragmentEnd - blockOffset;
                    blockOffset = fragmentEnd;
                    bool finishedTableFooter = false;
                    if (repeatTrailing) {
                        foreach (HtmlRenderVisual visual in trailingGroup!.Visuals) {
                            fragmentVisuals.Add(visual.Translate(0D, y, fragmentVisuals.Count));
                        }
                        y += trailingHeight;
                        if (blockOffset >= trailingGroup.ContentEndsAt - 0.0001D) {
                            blockOffset = trailingGroup.SourceEndsAt;
                            finishedTableFooter = true;
                        }
                    }

                    AddTranslatedVisuals(visuals, repeatContinuation || repeatTrailing
                        ? OrderRepeatedTableFragmentPaint(fragmentVisuals) : fragmentVisuals,
                        pageGeometry.Margins.Left, 0D, block);

                    if (blockOffset < block.Height - 0.0001D) {
                        // A wrapper can contain content after the completed table.
                        // Let that content use the space after its final footer.
                        if (finishedTableFooter
                            && ResolveForcedBreakAt(block.ForcedBreaks, blockOffset) == HtmlPageBreakTarget.None
                            && string.Equals(ResolvePageNameAt(block.ForcedBreaks, blockOffset, currentPageName), currentPageName, StringComparison.Ordinal)
                            && y < ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - 0.0001D) continue;
                        HtmlInlineBreakProgress? continuationProgress = ResolveInlineContinuationProgress(block, blockOffset);
                        string? nextPageName = ResolvePageNameAt(block.ForcedBreaks, blockOffset, currentPageName);
                        ExtendPagedBodyBackgroundThroughFragmentSlack(block, pageGeometry, visuals, y);
                        CommitPage(pages, visuals, pageGeometry, currentPageName);
                        BeginPage(nextPageName);
                        currentPageName = nextPageName;
                        pageWidth = pageGeometry.Width;
                        pageHeight = pageGeometry.Height;
                        contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                        HtmlPageBreakTarget internalBreak = ResolveForcedBreakAt(block.ForcedBreaks, blockOffset);
                        if (internalBreak != HtmlPageBreakTarget.None) {
                            EnsurePageSide(internalBreak, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                            pageWidth = pageGeometry.Width;
                            pageHeight = pageGeometry.Height;
                            contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
                        }
                        if (RequiresPageRelayout(block, pageGeometry)) {
                            if (continuationProgress.HasValue
                                && TryRelayoutInlineContinuation(block, pageGeometry, continuationProgress.Value, out HtmlRenderFlowBlock reflowed)) {
                                block = QualifySharedTableRepetition(reflowed.ForPagination());
                                blockOffset = 0D;
                            } else {
                                ReportPageContinuationReflowPending(block, pageGeometry);
                            }
                        } else if (internalBreak == HtmlPageBreakTarget.None) {
                            blockOffset = SkipUnpaintedLeadingMarginAtPageStart(block, blockOffset);
                        }
                    }
                }
            }

            HtmlPageBreakTarget breakAfter = ResolveForcedBreakAt(block.ForcedBreaks, block.Height);
            if (IsForcedPageBreak(block.BreakAfter)) breakAfter = block.BreakAfter;
            if (breakAfter != HtmlPageBreakTarget.None && index < blocks.Count - 1) {
                CommitPage(pages, visuals, pageGeometry, currentPageName);
                BeginPage(currentPageName);
                EnsurePageSide(breakAfter, pages, ref visuals, ref y, ref pageGeometry, currentPageName);
                pageWidth = pageGeometry.Width;
                pageHeight = pageGeometry.Height;
                contentHeight = ResolvePageBodyContentHeight(pages.Count + 1, pageGeometry);
            }
        }

        CommitPage(pages, visuals, pageGeometry, currentPageName);
        while (pages.Count < Math.Max(_footnotePlan.MaximumPageNumber, _pageFloatPlan.MaximumPageNumber)) {
            BeginPage(currentPageName);
            CommitPage(pages, visuals, pageGeometry, currentPageName);
        }
        return new HtmlRenderDocument(HtmlRenderMode.Paged, ApplyPageMarginContent(ProjectRelativePagedPaint(pages)), _diagnostics, _fonts, _metadata, _bookmarkDefinitions);
    }

}
