namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static bool IsSafeParallelItemBreak(
        HtmlRenderFlowBlock item,
        double offset,
        IDictionary<HtmlRenderFlowBlock, double> atomicVisualBottoms,
        IDictionary<HtmlRenderFlowBlock, IReadOnlyList<(double Top, double Bottom)>> atomicVisualRanges,
        double fullPageHeight = 0D) {
        if (offset <= 0.0001D || offset >= item.PagedPaintExtent - 0.0001D) return true;
        int precedingBreakIndex = UpperBound(item.BreakOffsets, offset + 0.0001D) - 1;
        if (precedingBreakIndex >= 0 && Math.Abs(item.BreakOffsets[precedingBreakIndex] - offset) <= 0.0001D) return true;
        // A stretched item can have a large painted but content-free tail. Its background
        // may fragment while the neighboring column supplies the actual page break.
        if (!atomicVisualBottoms.TryGetValue(item, out double bottom)) {
            bottom = LastAtomicParallelVisualBottom(item.Visuals);
            atomicVisualBottoms[item] = bottom;
        }
        if (offset >= bottom - 0.0001D) return true;

        // Neighboring parallel items need not share identical line-box heights. A
        // sibling's legal break can also split this item in its unpainted space
        // between line boxes or child blocks, provided its own break constraints
        // and atomic visuals remain intact.
        if ((item.AvoidBreakInside && (fullPageHeight <= 0D || item.PagedPaintExtent <= fullPageHeight + 0.0001D))
            || item.AvoidBreakRanges.Any(range =>
                offset > range.Start + 0.0001D && offset < range.End - 0.0001D
                && (fullPageHeight <= 0D || range.End - range.Start <= fullPageHeight + 0.0001D))) return false;
        if (!atomicVisualRanges.TryGetValue(item, out IReadOnlyList<(double Top, double Bottom)>? ranges)) {
            ranges = CollectAtomicParallelVisualRanges(item.Visuals);
            atomicVisualRanges[item] = ranges;
        }
        // Leading spacing may have no preceding content break. Let a sibling
        // break there before a fitting image/keep, instead of forcing a cut
        // through that content later on an otherwise empty page.
        if (ranges.Count > 0 && offset <= ranges[0].Top + 0.0001D) return true;
        int previousBreakIndex = UpperBound(item.BreakOffsets, offset - 0.0001D) - 1;
        if (previousBreakIndex < 0 || item.BreakOffsets[previousBreakIndex] <= 0.0001D) return false;
        return !CrossesAtomicParallelVisual(ranges, offset);
    }

    private static double LastAtomicParallelVisualBottom(IEnumerable<HtmlRenderVisual> visuals,
        double verticalTranslation = 0D, bool includePaintAndMetadata = false) {
        double bottom = 0D;
        foreach (HtmlRenderVisual visual in visuals) {
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                _ => null
            };
            if (children != null) {
                double childTranslation = verticalTranslation;
                if (visual is HtmlRenderEffectGroup effect && TryGetVerticalPaintTranslation(effect.Transform, out double effectTranslation)) {
                    childTranslation += effectTranslation;
                }
                bottom = Math.Max(bottom, LastAtomicParallelVisualBottom(children, childTranslation, includePaintAndMetadata));
            }
            else if ((includePaintAndMetadata && visual is not HtmlRenderLayoutBox)
                     || visual is HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderFormField
                         or HtmlRenderShape { IsAtomicReplacedPlaceholder: true })
                bottom = Math.Max(bottom, visual.LayoutY + visual.LayoutHeight + verticalTranslation);
        }
        return bottom;
    }

    private static IReadOnlyList<(double Top, double Bottom)> CollectAtomicParallelVisualRanges(IEnumerable<HtmlRenderVisual> visuals) {
        var ranges = new List<(double Top, double Bottom)>();
        AppendAtomicParallelVisualRanges(visuals, ranges);
        ranges.Sort((left, right) => left.Top.CompareTo(right.Top));
        var merged = new List<(double Top, double Bottom)>();
        foreach ((double top, double bottom) in ranges) {
            if (merged.Count > 0 && top < merged[merged.Count - 1].Bottom - 0.0001D) {
                (double previousTop, double previousBottom) = merged[merged.Count - 1];
                merged[merged.Count - 1] = (previousTop, Math.Max(previousBottom, bottom));
            } else {
                merged.Add((top, bottom));
            }
        }
        return merged;
    }

    private static void AppendAtomicParallelVisualRanges(IEnumerable<HtmlRenderVisual> visuals, ICollection<(double Top, double Bottom)> ranges,
        double verticalTranslation = 0D) {
        foreach (HtmlRenderVisual visual in visuals) {
            IReadOnlyList<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup group => group.Visuals,
                HtmlRenderEffectGroup group => group.Visuals,
                HtmlRenderLogicalTextGroup group => group.Visuals,
                HtmlRenderPathClipGroup group => group.Visuals,
                HtmlRenderLayoutRegion group => group.Visuals,
                HtmlRenderSemanticGroup group => group.Visuals,
                _ => null
            };
            if (children != null) {
                double childTranslation = verticalTranslation;
                if (visual is HtmlRenderEffectGroup effect && TryGetVerticalPaintTranslation(effect.Transform, out double effectTranslation)) {
                    childTranslation += effectTranslation;
                }
                AppendAtomicParallelVisualRanges(children, ranges, childTranslation);
            }
            else if (visual is HtmlRenderText or HtmlRenderImage or HtmlRenderDrawing or HtmlRenderFormField
                     or HtmlRenderShape { IsAtomicReplacedPlaceholder: true })
                ranges.Add((visual.LayoutY + verticalTranslation, visual.LayoutY + visual.LayoutHeight + verticalTranslation));
        }
    }

    private static bool CrossesAtomicParallelVisual(IReadOnlyList<(double Top, double Bottom)> ranges, double offset) {
        int low = 0;
        int high = ranges.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (ranges[middle].Top < offset - 0.0001D) low = middle + 1;
            else high = middle;
        }
        return low > 0 && ranges[low - 1].Bottom > offset + 0.0001D;
    }

}
