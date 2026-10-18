namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Keeps a block's final content with the next block's first substantive content.
    /// If that content belongs to an unbreakable descendant, the keep range extends
    /// through that descendant. Pagination may still split an oversized range.
    /// </summary>
    private void AppendKeepWithNextRange(
        IReadOnlyList<HtmlRenderFlowBlock> children,
        int childIndex,
        double childStart,
        ICollection<HtmlRenderAvoidBreakRange> ranges,
        double? previousFlowStart = null) {
        if (childIndex <= 0) return;

        HtmlRenderFlowBlock previous = children[childIndex - 1];
        HtmlRenderFlowBlock current = children[childIndex];
        HtmlFragmentationScope scope = ResolveAvoidBreakBetweenScope(previous, current);
        if (scope == HtmlFragmentationScope.None) return;
        double previousStart = previousFlowStart ?? Math.Max(0D, childStart - previous.NormalFlowAdvance);
        if ((scope & HtmlFragmentationScope.Page) != 0)
            Append(HtmlFragmentationScope.Page, previousStart + previous.PagedSourceEntryGap);
        if ((scope & HtmlFragmentationScope.Column) != 0)
            Append(HtmlFragmentationScope.Column,
                previousStart + FinalKeptContentStart(previous, HtmlFragmentationScope.Column));

        void Append(HtmlFragmentationScope target, double keepStart) {
            if (HasForcedBoundary(previous, current, target)) return;
            double firstContentBreak = FirstKeptContentExtent(current, target);
            if (HasForcedBreakBeforeKeptContent(current, firstContentBreak, target)) return;
            if (target == HtmlFragmentationScope.Page)
                firstContentBreak = Math.Max(0D, firstContentBreak - current.PagedSourceEntryGap);
            double keepEnd = childStart + firstContentBreak;
            if (keepEnd > keepStart + 0.0001D)
                ranges.Add(new HtmlRenderAvoidBreakRange(keepStart, keepEnd, Scope: target));
        }
    }

    private bool AvoidsBreakBetween(HtmlRenderFlowBlock previous, HtmlRenderFlowBlock current,
        HtmlFragmentationScope fragmentationScope = HtmlFragmentationScope.Page) =>
        (ResolveAvoidBreakBetweenScope(previous, current) & fragmentationScope) != 0;

    private HtmlFragmentationScope ResolveAvoidBreakBetweenScope(HtmlRenderFlowBlock previous, HtmlRenderFlowBlock current) {
        HtmlFragmentationScope scope = HtmlFragmentationScope.None;
        if (previous.OwnerElement != null
            && _layoutStyles.TryGetValue(previous.OwnerElement, out HtmlRenderBoxStyle? previousStyle))
            scope |= previousStyle.AvoidBreakAfterScope;
        if (current.OwnerElement != null
            && _layoutStyles.TryGetValue(current.OwnerElement, out HtmlRenderBoxStyle? currentStyle))
            scope |= currentStyle.AvoidBreakBeforeScope;
        return scope;
    }

    private static bool HasForcedBoundary(HtmlRenderFlowBlock previous, HtmlRenderFlowBlock current,
        HtmlFragmentationScope scope = HtmlFragmentationScope.Page) =>
        IsForcedBreakForScope(previous.BreakAfter, scope)
        || IsForcedBreakForScope(current.BreakBefore, scope)
        || !string.Equals(previous.PageName, current.PageName, StringComparison.Ordinal);

    private static double FirstKeptContentExtent(HtmlRenderFlowBlock current,
        HtmlFragmentationScope fragmentationScope = HtmlFragmentationScope.Page) {
        if ((current.AvoidBreakInsideScope & fragmentationScope) != 0) return current.Height;
        double firstContentBreak = current.InlineBreakProgress
            .Where(progress => progress.LogicalCharacters > 0 && progress.Offset > 0.0001D)
            .Select(progress => progress.Offset)
            .DefaultIfEmpty(0D)
            .Min();
        if (firstContentBreak <= 0.0001D) {
            // Images, vector drawings and fields have no logical characters.
            // An entry before their padding is not their first content end.
            IReadOnlyList<(double Top, double Bottom)> atomicContent = CollectAtomicParallelVisualRanges(current.Visuals);
            double firstAtomicEnd = atomicContent.Count == 0 ? 0D : atomicContent[0].Bottom;
            firstContentBreak = current.BreakOffsets
                .Where(offset => offset > 0.01D && offset >= firstAtomicEnd - 0.0001D)
                .DefaultIfEmpty(current.Height)
                .Min();
        }

        foreach (HtmlRenderAvoidBreakRange range in current.AvoidBreakRanges.OrderBy(range => range.Start)) {
            if ((range.Scope & fragmentationScope) == 0) continue;
            if (range.Start > firstContentBreak + 0.0001D || range.End <= firstContentBreak + 0.0001D) continue;
            firstContentBreak = range.End;
        }

        return Math.Min(current.Height, firstContentBreak);
    }

    /// <summary>
    /// Finds the last substantive source fragment that can travel with the next
    /// block without carrying an otherwise fragmentable preceding paragraph whole.
    /// </summary>
    private double FinalKeptContentStart(HtmlRenderFlowBlock previous, HtmlFragmentationScope fragmentationScope) {
        if ((previous.AvoidBreakInsideScope & fragmentationScope) != 0) return 0D;
        double contentEnd = previous.InlineBreakProgress
            .Where(progress => progress.LogicalCharacters > 0)
            .Select(progress => progress.Offset).DefaultIfEmpty(0D).Max();
        IReadOnlyList<(double Top, double Bottom)> atomicContent = CollectAtomicParallelVisualRanges(previous.Visuals);
        if (atomicContent.Count > 0) contentEnd = Math.Max(contentEnd, atomicContent[atomicContent.Count - 1].Bottom);
        if (contentEnd <= 0.0001D) contentEnd = previous.Height;

        double start = 0D;
        for (int index = previous.BreakOffsets.Count - 1; index >= 0; index--) {
            CheckCancellation();
            ChargeLayoutOperation("column final kept fragment");
            double candidate = previous.BreakOffsets[index];
            if (candidate >= contentEnd - 0.0001D || candidate <= 0.0001D) continue;
            if (!IsAllowedLineBreak(previous, 0D, candidate)) continue;
            start = candidate;
            break;
        }
        foreach (HtmlRenderAvoidBreakRange range in previous.AvoidBreakRanges.OrderByDescending(range => range.End)) {
            CheckCancellation();
            ChargeLayoutOperation("column final kept ranges");
            if ((range.Scope & fragmentationScope) != 0
                && range.Start < start - 0.0001D && range.End > start + 0.0001D)
                start = range.Start;
        }
        return Math.Max(0D, start);
    }

    private static bool HasForcedBreakBeforeKeptContent(HtmlRenderFlowBlock block, double keptExtent,
        HtmlFragmentationScope scope = HtmlFragmentationScope.Page) =>
        block.ForcedBreaks.Any(breakPoint => IsForcedBreakForScope(breakPoint.Target, scope)
            && breakPoint.Offset <= keptExtent + 0.0001D);

    /// <summary>
    /// A top-level block has no common flow block carrying sibling avoid ranges.
    /// Move it with the next block's first line when both fit on an empty page.
    /// </summary>
    private bool ShouldMoveTopLevelKeepWithNext(
        IReadOnlyList<HtmlRenderFlowBlock> blocks,
        int index,
        HtmlRenderFlowBlock current,
        double remainingHeight,
        HtmlCssPageGeometry currentGeometry,
        int nextPageNumber) {
        if (index >= blocks.Count - 1 || remainingHeight <= 0D) return false;
        HtmlRenderFlowBlock next = blocks[index + 1];
        if (!AvoidsBreakBetween(current, next) || HasForcedBoundary(current, next)) return false;
        HtmlCssPageGeometry nextGeometry = _pageRules.ResolveGeometry(nextPageNumber, current.PageName, _options);
        double fullPageHeight = ResolvePageBodyContentHeight(nextPageNumber, nextGeometry);
        HtmlRenderFlowBlock targetCurrent = current;
        HtmlRenderFlowBlock targetNext = next;
        if (!SamePageGeometry(currentGeometry, nextGeometry)) {
            SetActivePageGeometry(nextGeometry);
            try {
                targetCurrent = RelayoutTopLevelBlockForPage(current, nextGeometry);
                targetNext = RelayoutTopLevelBlockForPage(next, nextGeometry);
            } finally {
                SetActivePageGeometry(currentGeometry);
            }
        }
        double firstContentBreak = FirstKeptContentExtent(targetNext);
        if (HasForcedBreakBeforeKeptContent(targetNext, firstContentBreak)) return false;
        double keptHeight = targetCurrent.Height + firstContentBreak;
        return keptHeight <= fullPageHeight + 0.0001D
            && keptHeight > remainingHeight + 0.0001D;
    }
}
