namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
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
                if (block.BreakOffsets.Count <= 2
                    || (block.AvoidBreakInside && block.PagedPaintExtent <= fullPageHeight + 0.0001D)
                    || block.AvoidBreakRanges.Any(range =>
                        localOffset > range.Start + 0.0001D && localOffset < range.End - 0.0001D
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
}
