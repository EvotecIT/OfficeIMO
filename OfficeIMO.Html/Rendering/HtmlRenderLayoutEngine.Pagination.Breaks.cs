namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private double SkipUnpaintedLeadingMarginAtPageStart(HtmlRenderFlowBlock block, double start) {
        double discardableMargin = 0D;
        foreach (HtmlInlineBreakProgress progress in block.InlineBreakProgress) {
            if (Math.Abs(progress.Offset - start) > 0.0001D || !progress.IsBlockEntry
                || progress.PageStartDiscardableMargin <= 0.0001D) continue;
            discardableMargin = progress.PageStartDiscardableMargin;
            break;
        }
        if (discardableMargin <= 0.0001D) return start;

        double afterMargin = Math.Min(block.Height, start + discardableMargin);
        if (block.ForcedBreaks.Any(item => item.Offset > start + 0.0001D && item.Offset < afterMargin - 0.0001D)
            || block.RunningStringAssignments.Any(item => item.Offset >= start - 0.0001D && item.Offset < afterMargin - 0.0001D)) {
            return start;
        }

        // The container has moved to a new page because its first child could
        // not fit. Its leading margin is blank space from the previous page,
        // not a decoration to carry ahead of the child on the new page.
        return afterMargin;
    }

    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available, double? maximumEnd = null, double fullPageHeight = 0D) {
        double limit = Math.Min(maximumEnd ?? block.Height, Math.Min(block.Height, start + available));
        IReadOnlyList<double> offsets = block.BreakOffsets;
        for (int index = UpperBound(offsets, limit + 0.0001D) - 1; index >= 0; index--) {
            double offset = offsets[index];
            if (offset <= start + 0.0001D) break;
            if (IsAllowedLineBreak(block, start, offset)
                && !BreaksAvoidedRangeThatFitsPage(block, start, offset, fullPageHeight)) return offset;
        }

        return start;
    }

    // When this page's full body can fit the card, prefer its entry or end.
    // Otherwise retain interior offsets for a legal split.
    private static bool BreaksAvoidedRangeThatFitsPage(HtmlRenderFlowBlock block, double start, double candidate, double fullPageHeight) =>
        fullPageHeight > 0D && block.AvoidBreakRanges.Any(range =>
            start <= range.Start + 0.0001D
            && candidate > range.Start + 0.0001D
            && candidate < range.End - 0.0001D
            && range.End - range.Start <= fullPageHeight + 0.0001D);

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

    private static bool IsAllowedLineBreak(HtmlRenderFlowBlock block, double start, double candidate) {
        foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) {
            IReadOnlyList<double> offsets = group.Offsets;
            int candidateIndex = UpperBound(offsets, candidate + 0.0001D) - 1;
            if (candidateIndex < 0 || Math.Abs(offsets[candidateIndex] - candidate) > 0.0001D) continue;
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
