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
            if (breakBefore == HtmlPageBreakTarget.None) breakBefore = block.BreakBefore;
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

            if (block.Height <= ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y && !HasInternalForcedBreak(block)) {
                AddTranslatedVisuals(visuals, block.Visuals, pageGeometry.Margins.Left, y, block);
                RecordRunningStringAssignments(block, 0D, block.Height, y);
                y += block.Height;
            } else {
                double blockOffset = 0D;
                while (blockOffset < block.Height - 0.0001D) {
                    CheckCancellation();
                    double flexAvailable = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
                    if (!_pagedFlexAlignedBlocks.Contains(block)
                        && flexAvailable > 0.0001D
                        && block.Height > blockOffset + flexAvailable + 0.0001D
                        && TryRelayoutBlockForFlexPagination(block, blockOffset, flexAvailable, contentHeight,
                            pages.Count + 1, pageGeometry, out HtmlRenderFlowBlock alignedFlexBlock)) {
                        block = alignedFlexBlock.ForPagination();
                        _pagedFlexAlignedBlocks.Add(block);
                    }
                    HtmlRenderContinuationGroup? continuationGroup = block.ContinuationGroups.FirstOrDefault(group => group.AppliesAt(blockOffset));
                    bool repeatContinuation = blockOffset > 0.0001D && continuationGroup != null && continuationGroup.Visuals.Count > 0 && continuationGroup.Height > 0D;
                    double continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                    double rawAvailable = ResolvePageBodyBottom(pages.Count + 1, pageGeometry) - y;
                    HtmlRenderTrailingGroup? trailingGroup = ResolveTrailingGroup(block, blockOffset, Math.Max(0D, rawAvailable - continuationHeight), contentHeight, out double fragmentLimit);
                    bool repeatTrailing = trailingGroup != null && trailingGroup.Visuals.Count > 0 && trailingGroup.Height > 0D;
                    double trailingHeight = repeatTrailing ? trailingGroup!.Height : 0D;
                    double available = rawAvailable - continuationHeight - trailingHeight;
                    bool forcedBreakFits = TryGetNextForcedBreak(block.ForcedBreaks, blockOffset, out HtmlRenderForcedBreak? forcedBreak)
                        && forcedBreak!.Offset <= fragmentLimit + 0.0001D
                        && forcedBreak.Offset <= blockOffset + available + 0.0001D;
                    double fragmentEnd = forcedBreakFits
                        ? forcedBreak!.Offset
                        : available > 0.0001D
                            ? FindFragmentEnd(block, blockOffset, available, fragmentLimit, contentHeight)
                            : blockOffset;
                    if (fragmentEnd <= blockOffset + 0.0001D) {
                        if (y > ResolvePageBodyTop(pages.Count + 1, pageGeometry) + 0.0001D
                            || _pageFloatPlan.ReservedHeight(pages.Count + 1, "top") > 0D
                            || _pageFloatPlan.ReservedHeight(pages.Count + 1, "bottom") > 0D) {
                            CommitPage(pages, visuals, pageGeometry, currentPageName);
                            BeginPage(currentPageName);
                            continue;
                        }

                        bool originalContinuation = repeatContinuation;
                        bool originalTrailing = repeatTrailing;
                        bool foundFallback = false;
                        if (originalContinuation) {
                            double candidateAvailable = rawAvailable - trailingHeight;
                            double candidateEnd = candidateAvailable > 0.0001D
                                ? FindFragmentEnd(block, blockOffset, candidateAvailable, fragmentLimit, contentHeight)
                                : blockOffset;
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = false;
                                continuationHeight = 0D;
                                available = candidateAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (!foundFallback && originalTrailing) {
                            double candidateAvailable = rawAvailable - (originalContinuation ? continuationGroup!.Height : 0D);
                            double candidateEnd = candidateAvailable > 0.0001D
                                ? FindFragmentEnd(block, blockOffset, candidateAvailable, fragmentLimit, contentHeight)
                                : blockOffset;
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = originalContinuation;
                                continuationHeight = repeatContinuation ? continuationGroup!.Height : 0D;
                                repeatTrailing = false;
                                trailingHeight = 0D;
                                available = candidateAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (!foundFallback && originalContinuation && originalTrailing) {
                            double candidateEnd = FindFragmentEnd(block, blockOffset, rawAvailable, fragmentLimit, contentHeight);
                            if (candidateEnd > blockOffset + 0.0001D) {
                                repeatContinuation = false;
                                continuationHeight = 0D;
                                repeatTrailing = false;
                                trailingHeight = 0D;
                                available = rawAvailable;
                                fragmentEnd = candidateEnd;
                                foundFallback = true;
                            }
                        }

                        if (foundFallback) {
                            if (originalContinuation && !repeatContinuation) {
                                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableHeaderRepeatSuppressed, "A repeated table header was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
                            }

                            if (originalTrailing && !repeatTrailing) {
                                _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.TableFooterRepeatSuppressed, "A repeated table footer was suppressed because it left no safe body-row break on an empty page.", HtmlDiagnosticSeverity.Warning, block.Source);
                            }
                        } else {
                            repeatContinuation = false;
                            continuationHeight = 0D;
                            repeatTrailing = false;
                            trailingHeight = 0D;
                            available = Math.Max(0D, rawAvailable);
                            fragmentEnd = Math.Min(fragmentLimit, blockOffset + available);
                            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.ForcedFragment, "A layout block had no safe break opportunity within one page and was force-fragmented.", HtmlDiagnosticSeverity.Warning, block.Source);
                        }
                    }

                    if (repeatContinuation) {
                        AddTranslatedVisuals(visuals, continuationGroup!.Visuals, pageGeometry.Margins.Left, y, block);
                        y += continuationHeight;
                    }

                    IReadOnlyList<HtmlRenderVisual> fragment = SliceBlockVisuals(block, blockOffset, fragmentEnd,
                        fragmentEnd < block.Height - 0.0001D ? blockOffset + available : null);
                    AddTranslatedVisuals(visuals, fragment, pageGeometry.Margins.Left, y, block);
                    RecordRunningStringAssignments(block, blockOffset, fragmentEnd, y);
                    y += fragmentEnd - blockOffset;
                    blockOffset = fragmentEnd;
                    if (repeatTrailing) {
                        AddTranslatedVisuals(visuals, trailingGroup!.Visuals, pageGeometry.Margins.Left, y, block);
                        y += trailingHeight;
                        if (blockOffset >= trailingGroup.ContentEndsAt - 0.0001D) blockOffset = trailingGroup.SourceEndsAt;
                    }

                    if (blockOffset < block.Height - 0.0001D) {
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
                                block = reflowed.ForPagination();
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
            if (block.BreakAfter != HtmlPageBreakTarget.None) breakAfter = block.BreakAfter;
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
