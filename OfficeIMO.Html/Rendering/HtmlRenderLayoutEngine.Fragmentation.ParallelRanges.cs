using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>Imports known blank allocation tails left by projected column items.</summary>
    private IEnumerable<HtmlRenderBreakRange> EnumerateUnfragmentedAllocationBreakRanges(
        IEnumerable<FlowPaintLayer> layers, double extent, bool includeContainerTail = false) {
        double lastEnd = 0D;
        foreach (FlowPaintLayer layer in layers) {
            CheckCancellation();
            ChargeLayoutOperation("unfragmented allocation intervals");
            HtmlRenderFlowBlock block = layer.Block;
            foreach (HtmlRenderBreakRange range in block.SafeBreakRanges) yield return range.Translate(layer.Y);
            double paintedEnd = layer.Y + block.PagedPaintExtent;
            lastEnd = Math.Max(lastEnd, paintedEnd);
            if (block.UnfragmentedHeight is double allocated && allocated > block.PagedPaintExtent + 0.0001D) {
                yield return new HtmlRenderBreakRange(paintedEnd, Math.Min(extent, layer.Y + allocated));
            }
        }
        if (includeContainerTail && lastEnd < extent - 0.0001D)
            yield return new HtmlRenderBreakRange(lastEnd, extent);
    }

    /// <summary>
    /// Imports source intervals only where every overlapping item permits the
    /// same cut. Paragraph groups remain on the container for fragment-relative
    /// widow/orphan checks by the shared paginator.
    /// </summary>
    private IReadOnlyList<HtmlRenderBreakRange> CollectSafeParallelBreakRanges(
        IEnumerable<(HtmlRenderFlowBlock Block, double Y)> sourcePlacements,
        double extent,
        IEnumerable<HtmlRenderBreakRange>? sourceRanges = null,
        double fullPageHeight = 0D) {
        (HtmlRenderFlowBlock Block, double Y)[] placements = sourcePlacements.ToArray();
        var imported = new List<HtmlRenderBreakRange>();
        foreach (HtmlRenderBreakRange range in sourceRanges
            ?? placements.SelectMany(placement => placement.Block.SafeBreakRanges.Select(range => range.Translate(placement.Y)))) {
            CheckCancellation();
            ChargeLayoutOperation("parallel source break intervals");
            imported.Add(range);
        }
        IReadOnlyList<HtmlRenderBreakRange> result = HtmlRenderBreakRange.Normalize(imported, extent);
        if (result.Count == 0) return result;
        var safety = new Dictionary<HtmlRenderFlowBlock, IReadOnlyList<HtmlRenderBreakRange>>();
        foreach ((HtmlRenderFlowBlock block, double y) in placements) {
            CheckCancellation();
            ChargeLayoutOperation("parallel break interval ownership");
            ChargeLayoutOperations(result.Count, "parallel break interval overlap");
            if (!result.Any(range => range.End > y + 0.0001D
                && range.Start < y + block.PagedPaintExtent - 0.0001D)) continue;
            if (!safety.TryGetValue(block, out IReadOnlyList<HtmlRenderBreakRange>? allowed)) {
                allowed = ResolveParallelItemBreakRanges(block, fullPageHeight);
                safety.Add(block, allowed);
            }
            IEnumerable<HtmlRenderBreakRange> shared = allowed.Select(range => range.Translate(y));
            if (y >= 0D) shared = shared.Prepend(new HtmlRenderBreakRange(0D, Math.Min(extent, y)));
            double itemEnd = y + block.PagedPaintExtent;
            if (itemEnd <= extent) shared = shared.Append(new HtmlRenderBreakRange(Math.Max(0D, itemEnd), extent));
            result = HtmlRenderBreakRange.Intersect(result, HtmlRenderBreakRange.Normalize(shared, extent));
            if (result.Count == 0) break;
        }
        return result;
    }

    private IReadOnlyList<HtmlRenderBreakRange> ResolveParallelItemBreakRanges(HtmlRenderFlowBlock block, double fullPageHeight) {
        double extent = block.PagedPaintExtent;
        if (HasUnsupportedParallelRangeGeometry(block.Visuals))
            return new[] { new HtmlRenderBreakRange(0D, 0D), new HtmlRenderBreakRange(extent, extent) };

        IReadOnlyList<HtmlRenderBreakRange> allowed = new[] { new HtmlRenderBreakRange(0D, extent) };
        // A nested table's internal reservations cannot be inferred from blank
        // paint. Its owning layout must supply an explicit interval or endpoint.
        if (ContainsParallelRangeTable(block.Visuals)) {
            allowed = HtmlRenderBreakRange.Intersect(allowed,
                HtmlRenderBreakRange.Normalize(block.SafeBreakRanges.Concat(block.BreakOffsets.Select(
                    offset => new HtmlRenderBreakRange(offset, offset))), extent));
        }
        IReadOnlyList<(double Top, double Bottom)> atomicRanges = CollectAtomicParallelVisualRanges(block.Visuals);
        allowed = HtmlRenderBreakRange.Exclude(allowed,
            atomicRanges.Select(range => new HtmlRenderBreakRange(range.Top, range.Bottom)));
        IEnumerable<HtmlRenderBreakRange> keeps = block.AvoidBreakRanges
            .Where(range => fullPageHeight <= 0D || range.End - range.Start <= fullPageHeight + 0.0001D)
            .Select(range => new HtmlRenderBreakRange(range.Start, range.End));
        if (block.AvoidBreakInside && (fullPageHeight <= 0D || extent <= fullPageHeight + 0.0001D))
            keeps = keeps.Prepend(new HtmlRenderBreakRange(0D, extent));
        allowed = HtmlRenderBreakRange.Exclude(allowed, keeps);

        // Existing carrier safety changes at these source endpoints. Paragraph
        // constraints depend on the actual fragment start and stay with the
        // shared paginator instead of being evaluated as a first-page fragment.
        var boundaries = new SortedSet<double>(block.BreakOffsets);
        foreach (HtmlRenderBreakRange range in block.SafeBreakRanges) {
            ChargeLayoutOperation("parallel source interval boundaries");
            boundaries.Add(range.Start);
            boundaries.Add(range.End);
        }
        var atomicBottoms = new Dictionary<HtmlRenderFlowBlock, double>();
        var atomic = new Dictionary<HtmlRenderFlowBlock, IReadOnlyList<(double Top, double Bottom)>> { [block] = atomicRanges };
        var result = new List<HtmlRenderBreakRange>();
        double[] semanticBoundaries = boundaries.ToArray();
        foreach (HtmlRenderBreakRange range in allowed) {
            CheckCancellation();
            ChargeLayoutOperation("parallel safe interval qualification");
            var cuts = new List<double> { range.Start };
            for (int index = UpperBound(semanticBoundaries, range.Start + 0.0001D);
                index < semanticBoundaries.Length && semanticBoundaries[index] < range.End - 0.0001D; index++) {
                ChargeLayoutOperation("parallel interval semantic partition");
                cuts.Add(semanticBoundaries[index]);
            }
            if (range.End > range.Start) cuts.Add(range.End);
            for (int index = 0; index < cuts.Count; index++) {
                double cut = cuts[index];
                bool startSafe = IsSafe(cut);
                if (startSafe) result.Add(new HtmlRenderBreakRange(cut, cut));
                if (index + 1 >= cuts.Count || cuts[index + 1] <= cut + 0.0001D) continue;
                double end = cuts[index + 1];
                if (startSafe && IsSafe(end) && IsSafe(cut + (end - cut) / 2D))
                    result.Add(new HtmlRenderBreakRange(cut, end));
            }
        }
        return HtmlRenderBreakRange.Normalize(result, extent);

        bool IsSafe(double offset) {
            ChargeLayoutOperation("parallel interval content safety");
            return IsSafeParallelItemBreak(block, offset, atomicBottoms, atomic, fullPageHeight,
                keepsAlreadyQualified: true);
        }
    }

    private bool HasUnsupportedParallelRangeGeometry(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            CheckCancellation();
            ChargeLayoutOperation("parallel interval visual geometry");
            if (visual is HtmlRenderEffectGroup effect && !effect.Transform.Equals(OfficeTransform.Identity)) return true;
            if (GetGroupChildren(visual) is { } children && HasUnsupportedParallelRangeGeometry(children)) return true;
        }
        return false;
    }

    private bool ContainsParallelRangeTable(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            CheckCancellation();
            ChargeLayoutOperation("parallel interval table ownership");
            if (visual is HtmlRenderSemanticGroup table && IsTableFormattingGroup(table)) return true;
            if (GetGroupChildren(visual) is { } children && ContainsParallelRangeTable(children)) return true;
        }
        return false;
    }

    private IReadOnlyList<HtmlRenderBreakRange> ConstrainParallelPaintBreakRanges(
        IReadOnlyList<HtmlRenderBreakRange> ranges, IEnumerable<HtmlRenderVisual> visuals) {
        if (ranges.Count == 0) return ranges;
        if (HasUnsupportedParallelRangeGeometry(visuals)) return Array.Empty<HtmlRenderBreakRange>();
        return HtmlRenderBreakRange.Exclude(ranges, CollectAtomicParallelVisualRanges(visuals)
            .Select(range => new HtmlRenderBreakRange(range.Top, range.Bottom)));
    }
}
