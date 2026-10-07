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
        if (!AvoidsBreakBetween(previous, current) || HasForcedBoundary(previous, current)) return;

        double firstContentBreak = FirstKeptContentExtent(current);
        if (HasForcedBreakBeforeKeptContent(current, firstContentBreak)) return;
        double previousStart = previousFlowStart ?? Math.Max(0D, childStart - previous.Height);
        double keepEnd = childStart + firstContentBreak;
        if (keepEnd > previousStart + 0.0001D)
            ranges.Add(new HtmlRenderAvoidBreakRange(previousStart, keepEnd));
    }

    private bool AvoidsBreakBetween(HtmlRenderFlowBlock previous, HtmlRenderFlowBlock current) {
        bool avoidsAfter = previous.OwnerElement != null
            && _layoutStyles.TryGetValue(previous.OwnerElement, out HtmlRenderBoxStyle? previousStyle)
            && previousStyle.AvoidBreakAfter;
        bool avoidsBefore = current.OwnerElement != null
            && _layoutStyles.TryGetValue(current.OwnerElement, out HtmlRenderBoxStyle? currentStyle)
            && currentStyle.AvoidBreakBefore;
        return avoidsAfter || avoidsBefore;
    }

    private static bool HasForcedBoundary(HtmlRenderFlowBlock previous, HtmlRenderFlowBlock current) =>
        previous.BreakAfter != HtmlPageBreakTarget.None
        || current.BreakBefore != HtmlPageBreakTarget.None
        || !string.Equals(previous.PageName, current.PageName, StringComparison.Ordinal);

    private static double FirstKeptContentExtent(HtmlRenderFlowBlock current) {
        if (current.AvoidBreakInside) return current.Height;
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
            if (range.Start > firstContentBreak + 0.0001D || range.End <= firstContentBreak + 0.0001D) continue;
            firstContentBreak = range.End;
        }

        return Math.Min(current.Height, firstContentBreak);
    }

    private static bool HasForcedBreakBeforeKeptContent(HtmlRenderFlowBlock block, double keptExtent) =>
        block.ForcedBreaks.Any(breakPoint => breakPoint.Target != HtmlPageBreakTarget.None
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
