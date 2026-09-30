namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double SkipUnpaintedLeadingMarginAtPageStart(HtmlRenderFlowBlock block, double start) {
        double discardableMargin = ResolvePageStartDiscardableMargin(block, start);
        if (discardableMargin <= 0.0001D) return start;

        double afterMargin = Math.Min(block.Height, start + discardableMargin);
        if (block.ForcedBreaks.Any(item => item.Offset > start + 0.0001D && item.Offset <= afterMargin + 0.0001D)
            || block.RunningStringAssignments.Any(item => item.Offset >= start - 0.0001D && item.Offset < afterMargin - 0.0001D)
            || SliceBlockVisuals(block, start, afterMargin).Any(ContainsPageStartMarginContent)) {
            return start;
        }

        return afterMargin;
    }

    private static double ResolvePageStartDiscardableMargin(HtmlRenderFlowBlock block, double start) {
        double discardableMargin = block.HasCollapsibleMargins
            && block.CollapsibleMarginBottom > 0.0001D
            && Math.Abs(start - (block.Height - block.CollapsibleMarginBottom)) <= 0.0001D
                ? block.CollapsibleMarginBottom
                : 0D;
        foreach (HtmlInlineBreakProgress progress in block.InlineBreakProgress) {
            if (Math.Abs(progress.Offset - start) > 0.0001D || !(progress.IsBlockEntry || progress.IsBlockExit)
                || progress.PageStartDiscardableMargin <= 0.0001D) continue;
            discardableMargin = progress.PageStartDiscardableMargin;
            break;
        }
        return discardableMargin;
    }

    private static bool ContainsPageStartMarginContent(HtmlRenderVisual visual) {
        // Print-fitting boxes and empty wrappers measure layout but do not paint.
        // Retain every other leaf, including navigation and bookmark metadata.
        if (visual is HtmlRenderLayoutBox) return false;
        IReadOnlyList<HtmlRenderVisual>? children = GetGroupChildren(visual);
        return children == null || children.Any(ContainsPageStartMarginContent);
    }

    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available, double? maximumEnd = null, double fullPageHeight = 0D) {
        double limit = Math.Min(maximumEnd ?? block.Height, Math.Min(block.Height, start + available));
        IReadOnlyList<double> offsets = block.BreakOffsets;
        for (int index = UpperBound(offsets, limit + 0.0001D) - 1; index >= 0; index--) {
            double offset = offsets[index];
            if (offset <= start + 0.0001D) break;
            if (IsAllowedLineBreak(block, start, offset)
                && !BreaksAvoidedRangeThatFitsPage(block, start, offset, fullPageHeight, limit)) return offset;
        }

        return start;
    }

    // Keep authored ranges intact. A synthetic flex-row keep is weaker when
    // moving the whole row would leave more than half the page unused.
    private static bool BreaksAvoidedRangeThatFitsPage(HtmlRenderFlowBlock block, double start, double candidate, double fullPageHeight, double pageLimit) =>
        fullPageHeight > 0D && block.AvoidBreakRanges.Any(range =>
            start <= range.Start + 0.0001D
            && candidate > range.Start + 0.0001D
            && candidate < range.End - 0.0001D
            && range.End - range.Start <= fullPageHeight + 0.0001D
            && !(range.Soft && pageLimit - range.Start > fullPageHeight * 0.5D + 0.0001D));

    private static HtmlRenderTrailingGroup? ResolveTrailingGroup(HtmlRenderFlowBlock block, double start, double available, double fullPageHeight, out double fragmentLimit) {
        HtmlRenderTrailingGroup? active = block.TrailingGroups.FirstOrDefault(group => group.AppliesAt(start));
        if (active != null) {
            fragmentLimit = active.ContentEndsAt;
            return active;
        }

        HtmlRenderTrailingGroup? upcoming = block.TrailingGroups
            .Where(group => group.StartsAt > start + 0.0001D && group.StartsAt < start + available - 0.0001D)
            .OrderBy(group => group.StartsAt)
            .FirstOrDefault();
        if (upcoming == null) {
            fragmentLimit = block.Height;
            return null;
        }

        double candidateAvailable = Math.Max(0D, available - upcoming.Height);
        double candidateEnd = FindFragmentEnd(block, start, candidateAvailable, upcoming.ContentEndsAt, fullPageHeight);
        if (candidateEnd > upcoming.StartsAt + 0.0001D) {
            fragmentLimit = upcoming.ContentEndsAt;
            return upcoming;
        }

        fragmentLimit = upcoming.StartsAt;
        return null;
    }

    private static bool IsAllowedLineBreak(HtmlRenderFlowBlock block, double start, double candidate, bool checkInteriorBreaks = false) {
        foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
            // A final-line cut before a trailing paragraph margin leaves no widows.
            // Other final-line cuts must retain authored fixed-height flow.
            if (group.FinalLineMarginBreak && candidate >= group.End - 0.0001D
                && ResolvePageStartDiscardableMargin(block, candidate) > 0.0001D) continue;
            IReadOnlyList<double> offsets = group.Offsets;
            int candidateIndex = UpperBound(offsets, candidate + 0.0001D) - 1;
            bool exactLineBreak = candidateIndex >= 0 && Math.Abs(offsets[candidateIndex] - candidate) <= 0.0001D;
            // A flex sibling can supply a break in this item's unpainted line-box
            // space. Only flex rows need the wider paragraph-span check.
            if (!exactLineBreak && (!(group.CheckInteriorBreaks || checkInteriorBreaks)
                || candidate <= group.Start + 0.0001D || candidate >= group.End - 0.0001D)) continue;
            int firstFragmentLine = UpperBound(offsets, start + 0.0001D);
            int fragmentLines = candidateIndex >= firstFragmentLine ? candidateIndex - firstFragmentLine + 1 : 0;
            int remainingLines = offsets.Count - candidateIndex - 1 + (group.HasImplicitFinalLine ? 1 : 0);
            if (fragmentLines < group.Orphans || remainingLines < group.Widows) return false;
        }

        return true;
    }

    private static int UpperBound(IReadOnlyList<double> values, double target) {
        int low = 0;
        int high = values.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (values[middle] <= target) low = middle + 1;
            else high = middle;
        }

        return low;
    }

    private static bool HasInternalForcedBreak(HtmlRenderFlowBlock block) =>
        TryGetNextForcedBreak(block.ForcedBreaks, 0D, out HtmlRenderForcedBreak? forcedBreak)
        && forcedBreak!.Offset < block.Height - 0.0001D;

    private static bool TryGetNextForcedBreak(
        IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks,
        double offset,
        out HtmlRenderForcedBreak? forcedBreak) {
        int low = 0;
        int high = forcedBreaks.Count;
        double target = offset + 0.0001D;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (forcedBreaks[middle].Offset <= target) low = middle + 1;
            else high = middle;
        }

        forcedBreak = low < forcedBreaks.Count ? forcedBreaks[low] : null;
        return forcedBreak != null;
    }

    private static HtmlPageBreakTarget ResolveForcedBreakAt(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, double offset) {
        HtmlPageBreakTarget target = HtmlPageBreakTarget.None;
        foreach (HtmlRenderForcedBreak forcedBreak in forcedBreaks) {
            if (forcedBreak.Offset < offset - 0.0001D) continue;
            if (forcedBreak.Offset > offset + 0.0001D) break;
            target = forcedBreak.Target;
        }
        return target;
    }

    private static string? ResolvePageNameAt(IReadOnlyList<HtmlRenderForcedBreak> forcedBreaks, double offset, string? currentPageName) {
        string? pageName = currentPageName;
        foreach (HtmlRenderForcedBreak forcedBreak in forcedBreaks) {
            if (forcedBreak.Offset < offset - 0.0001D) continue;
            if (forcedBreak.Offset > offset + 0.0001D) break;
            if (forcedBreak.ChangesPageName) pageName = forcedBreak.PageName;
        }
        return pageName;
    }
}
