namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static void AppendFloatFragmentationMetadata(
        HtmlRenderFlowBlock block,
        InlineFloatPlacement placement,
        ICollection<HtmlRenderForcedBreak> forcedBreaks,
        ICollection<HtmlRenderLineBreakGroup> lineBreakGroups,
        ICollection<HtmlRenderContinuationGroup> continuationGroups,
        ICollection<HtmlRenderTrailingGroup> trailingGroups) {
        if (block.BreakBefore != HtmlPageBreakTarget.None) forcedBreaks.Add(new HtmlRenderForcedBreak(placement.Y, block.BreakBefore));
        foreach (HtmlRenderForcedBreak item in block.ForcedBreaks) forcedBreaks.Add(item.Translate(placement.Y));
        if (block.BreakAfter != HtmlPageBreakTarget.None) forcedBreaks.Add(new HtmlRenderForcedBreak(placement.Bottom, block.BreakAfter));
        foreach (HtmlRenderLineBreakGroup group in block.LineBreakGroups) lineBreakGroups.Add(group.Translate(placement.Y).WithInteriorBreaks());
        foreach (HtmlRenderContinuationGroup group in block.ContinuationGroups) continuationGroups.Add(group.Translate(placement.X, placement.Y));
        foreach (HtmlRenderTrailingGroup group in block.TrailingGroups) trailingGroups.Add(group.Translate(placement.X, placement.Y));
    }

    private IReadOnlyList<double> CollectSafeFloatBreaks(
        IEnumerable<double> existingBreaks,
        IReadOnlyList<InlineFloatPlacement> placements,
        IReadOnlyList<(double Top, double Bottom)> atomicVisualRanges,
        double fullPageHeight) {
        var breakOffsets = new SortedSet<double>(existingBreaks);
        foreach (InlineFloatPlacement placement in placements) {
            foreach (double offset in placement.Run.FloatingBlock!.BreakOffsets) {
                ChargeLayoutOperation("floated page break candidates");
                // Float entry/exit boundaries already belong to normal flow and
                // deferred-float relayout. Import only its internal opportunities.
                if (offset > 0.0001D && offset < placement.Height - 0.0001D) breakOffsets.Add(placement.Y + offset);
            }
        }

        var atomicBottoms = new Dictionary<HtmlRenderFlowBlock, double>();
        var atomicRanges = new Dictionary<HtmlRenderFlowBlock, IReadOnlyList<(double Top, double Bottom)>>();
        bool IsUnsafe(double offset) {
            ChargeLayoutOperation("floated page break safety");
            if (CrossesAtomicParallelVisual(atomicVisualRanges, offset)) return true;
            foreach (InlineFloatPlacement placement in placements) {
                ChargeLayoutOperation("parallel float break safety");
                if (offset <= placement.Y + 0.0001D || offset >= placement.Bottom - 0.0001D) continue;
                HtmlRenderFlowBlock block = placement.Run.FloatingBlock!;
                double localOffset = offset - placement.Y;
                // Keep an atomic float intact. A fragmentable float uses the same
                // safe-boundary contract as other side-by-side formatting items.
                if (block.BreakOffsets.Count <= 2 && !HasInteriorSafeBreakRanges(block)
                    || (block.AvoidBreakInside && block.PagedPaintExtent <= fullPageHeight + 0.0001D)
                    || block.AvoidBreakRanges.Any(range =>
                        (range.Scope & HtmlFragmentationScope.Page) != 0
                        && localOffset > range.Start + 0.0001D && localOffset < range.End - 0.0001D
                        && range.End - range.Start <= fullPageHeight + 0.0001D)
                    || !IsSafeParallelItemBreak(block, localOffset, atomicBottoms, atomicRanges, fullPageHeight)) return true;
            }
            return false;
        }

        breakOffsets.RemoveWhere(IsUnsafe);
        // These are container break opportunities, not paragraph character
        // progress. Inline continuation/reflow metadata remains with its owner.
        return breakOffsets.ToArray();
    }

    /// <summary>Qualifies imported float intervals against every crossing float and surrounding line carrier.</summary>
    private IReadOnlyList<HtmlRenderBreakRange> CollectSafeFloatBreakRanges(
        IEnumerable<HtmlRenderBreakRange> existingRanges,
        IReadOnlyList<InlineFloatPlacement> placements,
        IReadOnlyList<(double Top, double Bottom)> atomicVisualRanges,
        double fullPageHeight) {
        HtmlRenderBreakRange[] inherited = existingRanges.ToArray();
        double extent = Math.Max(inherited.Select(range => range.End).DefaultIfEmpty().Max(),
            placements.Select(placement => placement.Bottom).DefaultIfEmpty().Max());
        IEnumerable<HtmlRenderBreakRange> sources = inherited.Concat(placements.SelectMany(placement =>
            HtmlRenderBreakRange.Intersect(placement.Run.FloatingBlock!.SafeBreakRanges,
                new[] { new HtmlRenderBreakRange(0D, placement.Height) })
                .Select(range => range.Translate(placement.Y))));
        IReadOnlyList<HtmlRenderBreakRange> result = CollectSafeParallelBreakRanges(
            placements.Select(placement => (placement.Run.FloatingBlock!, placement.Y)), extent, sources, fullPageHeight);
        result = HtmlRenderBreakRange.Exclude(result, placements
            .Where(placement => placement.Run.FloatingBlock!.BreakOffsets.Count <= 2
                && !HasInteriorSafeBreakRanges(placement.Run.FloatingBlock!))
            .Select(placement => new HtmlRenderBreakRange(placement.Y, placement.Bottom)));
        return HtmlRenderBreakRange.Exclude(result,
            atomicVisualRanges.Select(range => new HtmlRenderBreakRange(range.Top, range.Bottom)));
    }

    private static bool HasInteriorSafeBreakRanges(HtmlRenderFlowBlock block) => block.SafeBreakRanges.Any(range =>
        range.End > 0.0001D && range.Start < block.PagedPaintExtent - 0.0001D);
}
