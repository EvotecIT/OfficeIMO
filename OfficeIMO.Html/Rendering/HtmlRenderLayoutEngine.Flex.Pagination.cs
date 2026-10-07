using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    // Record the space before a wrapped row separately from its content entry.
    // The paginator can discard an authored gap when the previous row ends a page.
    private static IEnumerable<HtmlInlineBreakProgress> ResolveWrappedFlexGapProgress(
        IReadOnlyList<FlexLine> lines, HtmlRenderBoxStyle style, double contentY, double rowGap) {
        if (style.FlexWrap != "wrap" || rowGap <= 0.0001D) yield break;
        for (int index = 1; index < lines.Count; index++) {
            FlexLine line = lines[index];
            double previousEnd = lines[index - 1].CrossOffset + lines[index - 1].CrossSize;
            // Distributed align-content space is not the authored row gap.
            if (Math.Abs(line.CrossOffset - previousEnd - rowGap) > 0.0001D
                || line.Items.Count == 0 || line.Items[0].Element == null) continue;
            yield return new HtmlInlineBreakProgress(contentY + previousEnd, 0,
                pageStartDiscardableMargin: rowGap, isFlexGap: true);
        }
    }

    private bool CanAlignPagedRowFlex(IElement element, HtmlRenderBoxStyle style, FlexLine line) {
        if (HasBoundedCrossSize(style)) return false;
        for (IElement? parent = element.ParentElement; parent != null; parent = parent.ParentElement) {
            if (_layoutStyles.TryGetValue(parent, out HtmlRenderBoxStyle? parentStyle)
                && HasBoundedCrossSize(parentStyle)) return false;
        }
        return line.Items.Count >= 2 && Math.Abs(line.CrossOffset) <= 0.0001D
            && line.Items.All(item => Math.Abs(item.CrossOffset) <= 0.0001D
            && !item.HasExplicitCrossSize && !item.Style.MaxHeight.HasValue && !item.Style.AspectRatio.HasValue
            && item.Block != null && item.Block.ForcedBreaks.Count == 0
            && item.Block.ContinuationGroups.Count == 0 && item.Block.TrailingGroups.Count == 0);
    }

    private static bool HasBoundedCrossSize(HtmlRenderBoxStyle style) =>
        style.ExplicitHeight.HasValue || style.MaxHeight.HasValue || style.AspectRatio.HasValue;

    // A single column has one vertical flow. Wrapped/reversed columns must not
    // lend one item's discardable gap to an independently positioned neighbor.
    private static IEnumerable<HtmlInlineBreakProgress> ResolveColumnFlexGapProgress(
        IReadOnlyList<FlexLine> lines, double contentY, double rowGap) {
        if (lines.Count != 1 || rowGap <= 0.0001D) yield break;
        IReadOnlyList<FlexItem> items = lines[0].Items;
        for (int index = 1; index < items.Count; index++) {
            FlexItem item = items[index];
            double previousEnd = items[index - 1].MainOffset + items[index - 1].Block!.Height;
            if (Math.Abs(item.MainOffset - previousEnd - rowGap) > 0.0001D) continue;
            yield return new HtmlInlineBreakProgress(contentY + previousEnd, 0,
                pageStartDiscardableMargin: rowGap, isFlexGap: true);
        }
    }

    // A normal single column has one vertical flow. Carry spacing and actual
    // line-end progress with no owner/resume target: this does not implement column
    // continuation reflow or borrow margins from wrapped/reversed neighbors.
    private static IEnumerable<HtmlInlineBreakProgress> ResolveColumnFlexBreakProgress(
        IReadOnlyList<FlexLine> lines, HtmlRenderBoxStyle style, double contentY) {
        if (lines.Count != 1 || style.FlexDirection != "column") yield break;
        foreach (FlexItem item in lines[0].Items) {
            HtmlRenderFlowBlock child = item.Block!;
            double childStart = contentY + item.MainOffset;
            double top = child.HasCollapsibleMargins
                ? Math.Max(0D, child.CollapsibleMarginTop - child.LeadingFlowAdjustment) : 0D;
            if (top > 0.0001D) {
                yield return new HtmlInlineBreakProgress(childStart, 0,
                    isBlockEntry: true, pageStartDiscardableMargin: top);
            }
            if (child.HasCollapsibleMargins && child.CollapsibleMarginBottom > 0.0001D
                && child.Height > child.CollapsibleMarginBottom + 0.0001D) {
                yield return new HtmlInlineBreakProgress(childStart + child.Height - child.CollapsibleMarginBottom, 0,
                    pageStartDiscardableMargin: child.CollapsibleMarginBottom, isBlockExit: true);
            }
            foreach (HtmlInlineBreakProgress progress in child.InlineBreakProgress) {
                if (progress.LogicalCharacters > 0 && progress.Offset > 0.0001D) {
                    // Keep-with-next needs the first substantive line end,
                    // rather than a column child's entry before its margin.
                    yield return new HtmlInlineBreakProgress(childStart + progress.Offset, progress.LogicalCharacters);
                }
                if (!(progress.IsBlockEntry || progress.IsBlockExit)
                    || progress.PageStartDiscardableMargin <= 0.0001D) continue;
                yield return new HtmlInlineBreakProgress(childStart + progress.Offset, 0,
                    isBlockEntry: progress.IsBlockEntry,
                    pageStartDiscardableMargin: progress.PageStartDiscardableMargin,
                    isBlockExit: progress.IsBlockExit);
            }
        }
    }

    private IReadOnlyList<HtmlRenderAvoidBreakRange> ResolveColumnFlexAvoidBreakRanges(
        IReadOnlyList<FlexLine> lines, HtmlRenderBoxStyle style, double contentY) {
        if (lines.Count != 1 || style.FlexDirection != "column") return Array.Empty<HtmlRenderAvoidBreakRange>();
        IReadOnlyList<FlexItem> items = lines[0].Items;
        HtmlRenderFlowBlock[] children = items.Select(item => item.Block!).ToArray();
        var ranges = new List<HtmlRenderAvoidBreakRange>();
        for (int index = 0; index < items.Count; index++) {
            double start = contentY + items[index].MainOffset;
            HtmlRenderFlowBlock child = children[index];
            if (child.AvoidBreakInside) ranges.Add(new HtmlRenderAvoidBreakRange(start, start + child.Height));
            ranges.AddRange(child.AvoidBreakRanges.Select(range => range.Translate(start)));
            AppendKeepWithNextRange(children, index, start, ranges,
                index > 0 ? contentY + items[index - 1].MainOffset : null);
        }
        return ranges;
    }

    // Margin truncation must observe the same direct and descendant forced
    // boundaries as a block flow. Parallel/reversed columns retain their
    // existing contract rather than borrowing a single flow's page changes.
    private static IEnumerable<HtmlRenderForcedBreak> ResolveColumnFlexForcedBreaks(
        IReadOnlyList<FlexLine> lines, HtmlRenderBoxStyle style, double contentY) {
        if (lines.Count != 1 || style.FlexDirection != "column") yield break;
        string? childPageName = lines[0].Items.FirstOrDefault()?.Block?.PageName;
        foreach (FlexItem item in lines[0].Items) {
            HtmlRenderFlowBlock child = item.Block!;
            double start = contentY + item.MainOffset;
            if (!string.Equals(childPageName, child.PageName, StringComparison.Ordinal)) {
                yield return new HtmlRenderForcedBreak(start, HtmlPageBreakTarget.Page, child.PageName, changesPageName: true);
            }
            childPageName = child.PageName;
            if (child.BreakBefore != HtmlPageBreakTarget.None)
                yield return new HtmlRenderForcedBreak(start, child.BreakBefore);
            foreach (HtmlRenderForcedBreak boundary in child.ForcedBreaks) yield return boundary.Translate(start);
            if (child.BreakAfter != HtmlPageBreakTarget.None)
                yield return new HtmlRenderForcedBreak(start + child.Height, child.BreakAfter);
        }
    }

    private bool TryRelayoutBlockForFlexPagination(
        HtmlRenderFlowBlock block,
        double blockOffset,
        double available,
        double pageHeight,
        int pageNumber,
        HtmlCssPageGeometry geometry,
        out HtmlRenderFlowBlock reflowed) {
        reflowed = block;
        if (block.OwnerElement == null || HasInternalForcedBreak(block) || _pagedRowFlexEligibleElements.Count == 0
            || !_pagedRowFlexEligibleElements.Any(element => ContainsElementOrSelf(block.OwnerElement, element))) return false;
        if (!HasPagedRowFlexAtBoundary(block.OwnerElement, blockOffset, available, pageHeight,
            out bool restoresTextLines)) return false;
        double remainingPages = Math.Ceiling(Math.Max(0D, block.Height - blockOffset - available) / pageHeight);
        if (remainingPages > 64D) return false;
        for (int page = 1; page <= (int)remainingPages; page++) {
            HtmlCssPageGeometry nextGeometry = _pageRules.ResolveGeometry(pageNumber + page, block.PageName, _options);
            if (!SamePageGeometry(geometry, nextGeometry)
                || Math.Abs(pageHeight - ResolvePageBodyContentHeight(pageNumber + page, nextGeometry)) > 0.0001D)
                return false;
        }
        double baselineEnd = FindFragmentEnd(block, blockOffset, available, fullPageHeight: pageHeight);
        // A full relayout is warranted for a visibly stranded region, not for
        // ordinary line-height slack at every page boundary of a long document.
        if (!restoresTextLines
            && baselineEnd >= blockOffset + available - Math.Max(16D, pageHeight * 0.15D)) return false;

        IElement root = _document.Body ?? _document.DocumentElement ?? block.OwnerElement;
        bool isRoot = ReferenceEquals(block.OwnerElement, root);
        if (!isRoot && !ReferenceEquals(block.OwnerElement.ParentElement, root)) return false;
        HtmlRenderBoxStyle rootStyle = _styleResolver.Resolve(root, geometry.ContentWidth);
        HtmlRenderBoxStyle style = isRoot
            ? rootStyle
            : _styleResolver.Resolve(block.OwnerElement, geometry.ContentWidth, rootStyle);
        double leadingAdjustment = isRoot ? 0D : block.LeadingFlowAdjustment;
        var boundary = new PagedFloatBoundary(
            blockOffset + available + leadingAdjustment,
            pageHeight,
            blockOffset + leadingAdjustment);
        _pagedFlexAlignedInRelayout = false;
        HtmlRenderFlowBlock candidate = isRoot
            ? LayoutRootElement(root, geometry.ContentWidth, rootStyle, pageBoundary: boundary)
            : LayoutElement(block.OwnerElement, geometry.ContentWidth, style, rootStyle, 1,
                pageBoundary: boundary).AdjustLeadingFlowSpace(leadingAdjustment);
        if (!_pagedFlexAlignedInRelayout) return false;
        double candidateEnd = FindFragmentEnd(candidate, blockOffset, available, fullPageHeight: pageHeight);
        if (candidateEnd <= baselineEnd + 0.0001D) return false;
        reflowed = candidate;
        return true;
    }

    private bool HasPagedRowFlexAtBoundary(IElement owner, double blockOffset, double available, double pageHeight,
        out bool restoresTextLines) {
        restoresTextLines = false;
        double boundary = blockOffset + available;
        foreach (KeyValuePair<IElement, HtmlRenderFlowBlock> entry in _pagedRowFlexBlocks) {
            if (!_pagedRowFlexEligibleElements.Contains(entry.Key) || !ContainsElementOrSelf(owner, entry.Key)) continue;
            if (!_layoutStyles.TryGetValue(entry.Key, out HtmlRenderBoxStyle? style)) continue;
            double start;
            if (ReferenceEquals(owner, entry.Key)) {
                start = 0D;
            } else if (TryResolveContentOrigin(entry.Key, owner, out PositionedPoint contentOrigin)) {
                start = contentOrigin.Y - style.MarginTop - style.BorderTopWidth - style.PaddingTop;
            } else {
                continue;
            }
            if (start >= boundary - 0.0001D || start + entry.Value.Height <= boundary + 0.0001D) continue;
            double rowOffset = Math.Max(0D, blockOffset - start);
            double rowAvailable = boundary - Math.Max(blockOffset, start);
            if (rowAvailable <= 0D) continue;
            double rowEnd = FindFragmentEnd(entry.Value, rowOffset, rowAvailable, fullPageHeight: pageHeight);
            double rowSlack = boundary - start - rowEnd;
            double rowContentY = style.MarginTop + style.BorderTopWidth + style.PaddingTop;
            if (!_pagedRowFlexLines.TryGetValue(entry.Key, out FlexLine? line)
                || !WouldAlignPagedRowFlexItems(line,
                    new PagedFloatBoundary(boundary - start, pageHeight, blockOffset - start)
                        .Shift(rowContentY + line.CrossOffset),
                    rowEnd - rowContentY - line.CrossOffset, out bool rowRestoresTextLines)
                || (rowSlack <= Math.Max(16D, pageHeight * 0.15D) && !rowRestoresTextLines)) continue;
            restoresTextLines = rowRestoresTextLines;
            return true;
        }
        return false;
    }

    private bool WouldAlignPagedRowFlexItems(FlexLine line, PagedFloatBoundary boundary,
        double currentRowEnd, out bool restoresTextLines) {
        restoresTextLines = false;
        double cursor = Math.Max(0D, boundary.FragmentStart);
        double available = boundary.RemainingHeight - cursor;
        if (available <= 0.0001D || available > boundary.PageHeight + 0.0001D
            || line.Items.Max(item => item.Block!.Height) <= cursor + available + 0.0001D) return false;
        var cuts = new double[line.Items.Count];
        for (int index = 0; index < line.Items.Count; index++) {
            HtmlRenderFlowBlock item = line.Items[index].Block!;
            if (item.Height <= cursor + 0.0001D) {
                cuts[index] = item.Height;
                continue;
            }
            cuts[index] = item.Height <= cursor + available + 0.0001D
                ? item.Height
                : FindFragmentEnd(item, cursor, available, fullPageHeight: boundary.PageHeight);
            if (cuts[index] <= cursor + 0.0001D) return false;
        }
        double sharedEnd = cuts.Max();
        if (sharedEnd > cursor + available + 0.0001D
            || !line.Items.Select((item, index) =>
                item.Block!.Height > cuts[index] + 0.0001D && sharedEnd > cuts[index] + 0.0001D).Any(value => value))
            return false;
        // A sidebar's next atomic image may need the following page while the
        // prose column still has multiple legal line cuts in the current one.
        restoresTextLines = line.Items.Select((item, index) =>
            cuts[index] >= sharedEnd - 0.0001D && item.Block!.LineBreakGroups.Any(group =>
                group.Offsets.Count(offset => offset > currentRowEnd + 0.0001D
                    && offset <= cuts[index] + 0.0001D) >= 2)).Any(value => value);
        return true;
    }

    private bool TryAlignPagedRowFlexItems(FlexLine line, PagedFloatBoundary boundary) {
        if (line.Items.Count < 2 || line.CrossOffset > 0.0001D
            || line.Items.Any(item => Math.Abs(item.CrossOffset) > 0.0001D
                || item.Block == null
                || item.Block.ForcedBreaks.Count > 0
                || item.Block.ContinuationGroups.Count > 0
                || item.Block.TrailingGroups.Count > 0)) return false;

        double cursor = Math.Max(0D, boundary.FragmentStart);
        double available = boundary.RemainingHeight - cursor;
        if (available <= 0D && cursor <= 0.0001D) {
            available += (Math.Floor(-available / boundary.PageHeight) + 1D) * boundary.PageHeight;
        }
        if (available <= 0.0001D || available > boundary.PageHeight + 0.0001D) return false;
        bool changed = false;
        for (int page = 0; page < 64; page++) {
            CheckCancellation();
            double maximumHeight = line.Items.Max(item => item.Block!.Height);
            if (maximumHeight <= cursor + available + 0.0001D) break;

            var cuts = new double[line.Items.Count];
            bool canAdvance = true;
            for (int index = 0; index < line.Items.Count; index++) {
                HtmlRenderFlowBlock item = line.Items[index].Block!;
                if (item.Height <= cursor + 0.0001D) {
                    cuts[index] = item.Height;
                } else if (item.Height <= cursor + available + 0.0001D) {
                    cuts[index] = item.Height;
                } else {
                    cuts[index] = FindFragmentEnd(item, cursor, available, fullPageHeight: boundary.PageHeight);
                    if (cuts[index] <= cursor + 0.0001D) canAdvance = false;
                }
            }
            if (!canAdvance) break;

            // A completed column can contribute its far endpoint while another
            // column still needs an early cut. Do not stretch that sibling through
            // a large empty region solely to match the completed column's end.
            bool completedItem = line.Items.Any(item => item.Block!.Height <= cursor + available + 0.0001D);
            bool continuingItem = line.Items.Any(item => item.Block!.Height > cursor + available + 0.0001D);
            if (completedItem && continuingItem
                && cuts.Max() - cuts.Min() > Math.Max(16D, boundary.PageHeight * 0.25D)
                && HasUnalignedSharedFlexRowBreak(line, cursor, available, boundary.PageHeight)) break;

            double sharedEnd = cuts.Max();
            if (sharedEnd <= cursor + 0.0001D || sharedEnd > cursor + available + 0.0001D) break;
            for (int index = 0; index < line.Items.Count; index++) {
                HtmlRenderFlowBlock item = line.Items[index].Block!;
                double gap = sharedEnd - cuts[index];
                // A stretched background can outlive its last text or image.
                // Do not shift a large content-free tail and strand the next block.
                if (gap <= 0.0001D || item.Height <= cuts[index] + 0.0001D) continue;
                if (gap > Math.Max(16D, boundary.PageHeight * 0.25D)
                    && LastAtomicParallelVisualBottom(item.Visuals) <= cuts[index] + 0.0001D) continue;
                // Auto-height stretch may leave a content-free tail after the
                // last sidebar image. Use that tail for the inserted page-break
                // space instead of growing the row and moving its next sibling.
                double retainedEnd = new[] {
                    LastAtomicParallelVisualBottom(item.Visuals, includePaintAndMetadata: true),
                    item.LineBreakGroups.Select(group => group.End).DefaultIfEmpty().Max(),
                    item.BreakOffsets.Where(offset => offset < item.Height - 0.0001D).DefaultIfEmpty().Max(),
                    item.RunningStringAssignments.Select(assignment => assignment.Offset).DefaultIfEmpty().Max(),
                    item.InlineBreakProgress.Select(progress => progress.Offset).DefaultIfEmpty().Max(),
                    item.AvoidBreakRanges.Select(range => range.End).DefaultIfEmpty().Max()
                }.Max();
                bool absorbInStretchTail = !line.Items[index].HasExplicitCrossSize
                    && line.Items[index].Style.ExplicitHeight.HasValue
                    && retainedEnd + gap <= item.Height + 0.0001D;
                line.Items[index].Block = InsertFlexItemBreakGap(item, cuts[index], gap, absorbInStretchTail);
                changed = true;
            }

            cursor = sharedEnd;
            available = boundary.PageHeight;
        }

        if (changed) line.CrossSize = line.Items.Max(item => item.Block!.Height);
        return changed;
    }

    private static bool HasUnalignedSharedFlexRowBreak(FlexLine line, double cursor, double available, double pageHeight) {
        var atomicVisualBottoms = new Dictionary<HtmlRenderFlowBlock, double>();
        var atomicVisualRanges = new Dictionary<HtmlRenderFlowBlock, IReadOnlyList<(double Top, double Bottom)>>();
        return line.Items.SelectMany(item => item.Block!.BreakOffsets)
            .Where(offset => offset > cursor + 0.0001D && offset <= cursor + available + 0.0001D)
            .Distinct()
            .Any(offset => line.Items.All(item => {
                HtmlRenderFlowBlock block = item.Block!;
                return IsSafeParallelItemBreak(block, offset, atomicVisualBottoms, atomicVisualRanges)
                    && IsAllowedLineBreak(block, cursor, offset, checkInteriorBreaks: true)
                    && !(block.AvoidBreakInside && block.Height <= pageHeight + 0.0001D
                        && offset > 0.0001D && offset < block.Height - 0.0001D)
                    && !BreaksAvoidedRangeThatFitsPage(block, cursor, offset, pageHeight, cursor + available);
            }));
    }

    private HtmlRenderFlowBlock InsertFlexItemBreakGap(HtmlRenderFlowBlock block, double cut, double gap,
        bool absorbInStretchTail) {
        IReadOnlyList<HtmlRenderVisual> before = SliceVisuals(block.Visuals, 0D, cut);
        IReadOnlyList<HtmlRenderVisual> after = SliceVisuals(block.Visuals, cut, block.Height, includeEndAnchor: true);
        List<HtmlRenderVisual> visuals = before
            .Concat(after.Select((visual, index) => visual.Translate(0D, cut + gap, before.Count + index)))
            .ToList();
        double height = absorbInStretchTail ? block.Height : block.Height + gap;
        if (absorbInStretchTail) visuals = SliceVisuals(visuals, 0D, height, includeEndAnchor: true).ToList();
        double Shift(double offset) => offset >= cut - 0.0001D ? offset + gap : offset;

        return new HtmlRenderFlowBlock(
            block.Width,
            height,
            visuals,
            block.BreakBefore,
            block.BreakAfter,
            block.AvoidBreakInside,
            block.Source,
            block.BreakOffsets.Select(Shift).Concat(new[] { cut, cut + gap }),
            lineBreakGroups: block.LineBreakGroups.Select(group => new HtmlRenderLineBreakGroup(
                group.Offsets.Select(Shift),
                group.Orphans,
                group.Widows,
                group.HasImplicitFinalLine,
                group.Start >= cut - 0.0001D ? group.Start + gap : group.Start,
                group.End > cut + 0.0001D ? group.End + gap : group.End,
                group.CheckInteriorBreaks,
                group.FinalLineMarginBreak)),
            pageName: block.PageName,
            stackingZIndex: block.StackingZIndex,
            stackingSourceOrder: block.StackingSourceOrder,
            hasCollapsibleMargins: block.HasCollapsibleMargins,
            collapsibleMarginTop: block.CollapsibleMarginTop,
            collapsibleMarginBottom: block.CollapsibleMarginBottom,
            ownerElement: block.OwnerElement,
            collapsesThrough: block.CollapsesThrough,
            unclampedHeight: absorbInStretchTail ? block.UnclampedHeight : block.UnclampedHeight + gap,
            runningStringAssignments: block.RunningStringAssignments.Select(assignment =>
                assignment.Offset >= cut - 0.0001D ? assignment.Translate(gap) : assignment),
            inlineBreakProgress: block.InlineBreakProgress.Select(progress => new HtmlInlineBreakProgress(
                Shift(progress.Offset), progress.LogicalCharacters, progress.OwnerElement,
                progress.IsBlockEntry, progress.PageStartDiscardableMargin, progress.IsBlockExit, progress.IsFlexGap)),
            inlineContinuationStart: block.InlineContinuationStart,
            supportsInlineContinuationReflow: block.SupportsInlineContinuationReflow,
            layoutViewportWidth: block.LayoutViewportWidth,
            layoutViewportHeight: block.LayoutViewportHeight,
            leadingFlowAdjustment: block.LeadingFlowAdjustment,
            collapsibleMarginTopGroup: block.CollapsibleMarginTopGroup,
            collapsibleMarginBottomGroup: block.CollapsibleMarginBottomGroup,
            avoidBreakRanges: block.AvoidBreakRanges.Select(range => new HtmlRenderAvoidBreakRange(
                range.Start >= cut - 0.0001D ? range.Start + gap : range.Start,
                range.End > cut + 0.0001D ? range.End + gap : range.End,
                range.Soft)));
    }
}
