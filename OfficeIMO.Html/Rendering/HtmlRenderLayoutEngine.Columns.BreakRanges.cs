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

    /// <summary>Retains paragraph constraints for the physical lines present in each column fragment.</summary>
    private IEnumerable<HtmlRenderLineBreakGroup> EnumerateMultiColumnLineBreakGroups(MultiColumnPlan plan, double offsetY) {
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column paragraph projection");
            foreach (HtmlRenderLineBreakGroup group in SliceColumnLineBreakGroups(
                         fragment.Block.LineBreakGroups, fragment.Start, fragment.End, offsetY + fragment.Y)) {
                yield return group;
            }
        }
    }

    /// <summary>
    /// Clips a paragraph to actual source-fragment lines. A final included line
    /// becomes the fragment's implicit final line, so later columns cannot
    /// inflate its widow count or prevent finishing its complete content.
    /// </summary>
    private IEnumerable<HtmlRenderLineBreakGroup> SliceColumnLineBreakGroups(
        IEnumerable<HtmlRenderLineBreakGroup> groups, double start, double end, double offsetY) {
        const double tolerance = 0.0001D;
        if (end <= start + tolerance) yield break;
        double translation = offsetY - start;
        foreach (HtmlRenderLineBreakGroup group in groups) {
            CheckCancellation();
            ChargeLayoutOperations((long)group.Offsets.Count + 1L, "column paragraph line projection");
            if (group.End <= start + tolerance || group.Start >= end - tolerance) continue;
            if (group.Start >= start && group.End <= end) {
                yield return group.WithInteriorBreaks().Translate(translation);
                continue;
            }

            var presentOffsets = group.Offsets.Where(value => value > start + tolerance && value <= end + tolerance)
                .Select(value => Math.Min(end, value)).ToList();
            bool hasFinalLine = group.HasImplicitFinalLine;
            double groupEnd = Math.Min(group.End, end);
            bool hasOriginalEnd = group.End <= end + tolerance;
            if (!hasOriginalEnd) {
                if (presentOffsets.Count == 0) continue;
                groupEnd = presentOffsets[presentOffsets.Count - 1];
                presentOffsets.RemoveAt(presentOffsets.Count - 1);
                hasFinalLine = true;
            }
            yield return new HtmlRenderLineBreakGroup(
                presentOffsets.Select(value => value + translation), group.Orphans, group.Widows,
                hasFinalLine, Math.Max(group.Start, start) + translation, groupEnd + translation,
                checkInteriorBreaks: true, finalLineMarginBreak: hasOriginalEnd && group.FinalLineMarginBreak);
        }
    }

    /// <summary>Preserves source keeps that remain whole in an actual column fragment.</summary>
    private IEnumerable<HtmlRenderAvoidBreakRange> EnumerateMultiColumnAvoidBreakRanges(MultiColumnPlan plan, double offsetY) {
        foreach (MultiColumnFragment fragment in plan.Fragments) {
            CheckCancellation();
            ChargeLayoutOperation("column keep projection");
            foreach (HtmlRenderAvoidBreakRange range in SliceColumnAvoidBreakRanges(
                         fragment.Block, fragment.Start, fragment.End, offsetY + fragment.Y)) {
                yield return range;
            }
        }
    }

    /// <summary>
    /// Projects whole keeps only. A keep already split by the source column
    /// planner must not become a smaller, newly fitting keep on the outer page.
    /// </summary>
    private static IEnumerable<HtmlRenderAvoidBreakRange> SliceColumnAvoidBreakRanges(
        HtmlRenderFlowBlock block, double start, double end, double offsetY) {
        const double tolerance = 0.0001D;
        double translation = offsetY - start;
        if ((block.AvoidBreakInsideScope & HtmlFragmentationScope.Page) != 0
            && start <= tolerance && end >= block.Height - tolerance) {
            yield return new HtmlRenderAvoidBreakRange(Math.Max(start, 0D) + translation,
                Math.Min(end, block.Height) + translation, Scope: HtmlFragmentationScope.Page);
        }
        foreach (HtmlRenderAvoidBreakRange range in block.AvoidBreakRanges) {
            if ((range.Scope & HtmlFragmentationScope.Page) == 0) continue;
            if (range.Start < start - tolerance || range.End > end + tolerance) continue;
            yield return new HtmlRenderAvoidBreakRange(
                Math.Max(start, range.Start) + translation, Math.Min(end, range.End) + translation, range.Soft, HtmlFragmentationScope.Page);
        }
    }

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
