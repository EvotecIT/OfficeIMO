namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static double FindFragmentEnd(HtmlRenderFlowBlock block, double start, double available, double? maximumEnd = null) {
        double limit = Math.Min(maximumEnd ?? block.Height, Math.Min(block.Height, start + available));
        IReadOnlyList<double> offsets = block.BreakOffsets;
        for (int index = UpperBound(offsets, limit + 0.0001D) - 1; index >= 0; index--) {
            double offset = offsets[index];
            if (offset <= start + 0.0001D) break;
            if (IsAllowedLineBreak(block, start, offset)) return offset;
        }

        return start;
    }

    private static HtmlRenderTrailingGroup? ResolveTrailingGroup(HtmlRenderFlowBlock block, double start, double available, out double fragmentLimit) {
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
        double candidateEnd = FindFragmentEnd(block, start, candidateAvailable, upcoming.ContentEndsAt);
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
            int remainingLines = offsets.Count - candidateIndex;
            return fragmentLines >= group.Orphans && remainingLines >= group.Widows;
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
}
