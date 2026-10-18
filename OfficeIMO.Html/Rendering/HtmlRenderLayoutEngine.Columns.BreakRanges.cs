namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Projects imported source intervals into the column set, then retains only
    /// cuts permitted by every occupied fragment that crosses the same height.
    /// Empty column space imposes no additional constraint.
    /// </summary>
    private IReadOnlyList<HtmlRenderBreakRange> ResolveMultiColumnSafeBreakRanges(MultiColumnPlan plan) {
        IReadOnlyList<HtmlRenderBreakRange> candidates = HtmlRenderBreakRange.Normalize(
            plan.Fragments.SelectMany(fragment => SliceColumnSafeBreakRanges(
                fragment.Block.SafeBreakRanges, fragment.Start, fragment.End)
                .Select(range => range.Translate(fragment.Y))), plan.UsedHeight);
        if (candidates.Count == 0) return candidates;

        var domain = new[] { new HtmlRenderBreakRange(0D, plan.UsedHeight) };
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column safe break intersection");
            // Source-only anchors do not occupy a vertical content interval.
            if (fragment.End <= fragment.Start) continue;
            double end = fragment.Y + fragment.End - fragment.Start;
            var occupied = new[] { new HtmlRenderBreakRange(fragment.Y, end) };
            IEnumerable<HtmlRenderBreakRange> fragmentCuts = SliceColumnSafeBreakRanges(
                fragment.Block.SafeBreakRanges, fragment.Start, fragment.End)
                .Select(range => range.Translate(fragment.Y))
                .Concat(fragment.Block.BreakOffsets
                    .Where(offset => offset >= fragment.Start && offset <= fragment.End)
                    .Select(offset => new HtmlRenderBreakRange(
                        fragment.Y + offset - fragment.Start, fragment.Y + offset - fragment.Start)))
                .Concat(new[] { new HtmlRenderBreakRange(fragment.Y, fragment.Y), new HtmlRenderBreakRange(end, end) });
            candidates = HtmlRenderBreakRange.Intersect(candidates,
                HtmlRenderBreakRange.Exclude(domain, occupied).Concat(fragmentCuts));
            if (candidates.Count == 0) break;
        }
        return candidates;
    }

    /// <summary>Clips source intervals to an actual column or note fragment and makes them fragment-relative.</summary>
    private static IEnumerable<HtmlRenderBreakRange> SliceColumnSafeBreakRanges(
        IEnumerable<HtmlRenderBreakRange> ranges, double start, double end) =>
        HtmlRenderBreakRange.Intersect(ranges, new[] { new HtmlRenderBreakRange(start, end) })
            .Select(range => range.Translate(-start));

    /// <summary>
    /// Preserves the wrapper's final-break projection without widening an
    /// imported interval through trailing padding, margin, or marker geometry.
    /// </summary>
    private static IEnumerable<HtmlRenderBreakRange> ProjectColumnTerminalSafeBreakRanges(
        IEnumerable<HtmlRenderBreakRange> ranges, double sourceEnd, double offset, double outerEnd) {
        bool projectTerminal = Math.Abs(offset + sourceEnd - outerEnd) > 0.0001D;
        foreach (HtmlRenderBreakRange range in HtmlRenderBreakRange.Normalize(ranges, sourceEnd)) {
            if (!projectTerminal || range.End < sourceEnd - 0.0001D) {
                yield return range.Translate(offset);
                continue;
            }
            // Closed intervals cannot represent an open terminal endpoint.
            // Stay beyond the fragment lookup tolerance before projecting it.
            double interiorEnd = Math.Min(range.End, sourceEnd - 0.0002D);
            if (range.Start <= interiorEnd) yield return new HtmlRenderBreakRange(offset + range.Start, offset + interiorEnd);
            if (range.End >= sourceEnd) yield return new HtmlRenderBreakRange(outerEnd, outerEnd);
        }
    }
}
