using AngleSharp.Dom;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyList<HtmlRenderFlowBlock> BuildMultiColumnChildBlocks(
        IElement element,
        IEnumerable<INode> nodes,
        double width,
        HtmlRenderBoxStyle style,
        int depth,
        bool includeGeneratedBefore = true,
        bool includeGeneratedAfter = true) {
        if (!ContainsFloatingDescendant(element, width, style, depth)) {
            return BuildChildBlocks(element, nodes, width, style, depth,
                includeGeneratedBefore, includeGeneratedAfter);
        }

        // Share the ordinary float exclusions between sibling paragraphs. The
        // float does not advance normal flow, but its paint must survive slicing
        // and remain whole at a column boundary.
        var exclusions = new List<HtmlFloatExclusion>();
        IReadOnlyList<HtmlRenderFlowBlock> children = BuildChildBlocks(
            element, nodes, width, style, depth, includeGeneratedBefore, includeGeneratedAfter,
            emittedFloats: exclusions);
        if (exclusions.Count == 0) return children;

        var visuals = new List<HtmlRenderVisual>();
        var layers = new List<FlowPaintLayer>();
        var offsets = new SortedSet<double>();
        var lineGroups = new List<HtmlRenderLineBreakGroup>();
        var keeps = new List<HtmlRenderAvoidBreakRange>();
        var runningStrings = new List<HtmlCssRunningStringAssignment>();
        var progress = new List<HtmlInlineBreakProgress>();
        var forcedBreaks = new List<HtmlRenderForcedBreak>();
        var continuationGroups = new List<HtmlRenderContinuationGroup>();
        var trailingGroups = new List<HtmlRenderTrailingGroup>();
        double height = 0D;
        for (int index = 0; index < children.Count; index++) {
            HtmlRenderFlowBlock child = children[index];
            layers.Add(new FlowPaintLayer(child, 0D, height, index));
            foreach (double offset in child.BreakOffsets) offsets.Add(height + offset);
            lineGroups.AddRange(child.LineBreakGroups.Select(group => group.Translate(height)));
            AppendKeepWithNextRange(children, index, height, keeps);
            if (child.AvoidBreakInside) keeps.Add(new HtmlRenderAvoidBreakRange(height, height + child.Height));
            keeps.AddRange(child.AvoidBreakRanges.Select(range => range.Translate(height)));
            runningStrings.AddRange(child.RunningStringAssignments.Select(item => item.Translate(height)));
            progress.AddRange(child.InlineBreakProgress.Select(item => new HtmlInlineBreakProgress(
                height + item.Offset, item.LogicalCharacters, item.OwnerElement,
                item.IsBlockEntry, item.PageStartDiscardableMargin, item.IsBlockExit)));
            if (child.BreakBefore != HtmlPageBreakTarget.None) {
                forcedBreaks.Add(new HtmlRenderForcedBreak(height, child.BreakBefore));
            }
            forcedBreaks.AddRange(child.ForcedBreaks.Select(item => item.Translate(height)));
            continuationGroups.AddRange(child.ContinuationGroups.Select(group => group.Translate(0D, height)));
            trailingGroups.AddRange(child.TrailingGroups.Select(group => group.Translate(0D, height)));
            height += child.Height;
            if (child.BreakAfter != HtmlPageBreakTarget.None) {
                forcedBreaks.Add(new HtmlRenderForcedBreak(height, child.BreakAfter));
            }
        }
        AppendFlowPaintLayers(visuals, layers);
        height = Math.Max(height, exclusions.Max(item => item.Bottom));
        RemoveColumnBreaksInsideFloats(offsets, exclusions);
        return new[] { new HtmlRenderFlowBlock(
            width, height, visuals, HtmlPageBreakTarget.None, HtmlPageBreakTarget.None,
            false, HtmlRenderStyleResolver.DescribeSource(element), offsets,
            lineBreakGroups: lineGroups, continuationGroups: continuationGroups,
            trailingGroups: trailingGroups, runningStringAssignments: runningStrings,
            inlineBreakProgress: progress, forcedBreaks: forcedBreaks, avoidBreakRanges: keeps) };
    }

    private void RemoveColumnBreaksInsideFloats(SortedSet<double> offsets, IReadOnlyList<HtmlFloatExclusion> exclusions) {
        // A float edge is not itself a legal text break. Filter the existing
        // content cuts with a merged interval sweep instead of scanning every
        // historical float for each cut.
        CheckCancellation();
        var intervals = new List<(double Start, double End)>();
        foreach (HtmlFloatExclusion exclusion in exclusions.OrderBy(item => item.Y)) {
            CheckCancellation();
            ChargeLayoutOperation("column float interval");
            if (intervals.Count > 0 && exclusion.Y < intervals[intervals.Count - 1].End - 0.0001D) {
                var previous = intervals[intervals.Count - 1];
                intervals[intervals.Count - 1] = (previous.Start, Math.Max(previous.End, exclusion.Bottom));
            } else {
                intervals.Add((exclusion.Y, exclusion.Bottom));
            }
        }
        int interval = 0;
        foreach (double offset in offsets.ToArray()) {
            CheckCancellation();
            ChargeLayoutOperation("column float break filtering");
            while (interval < intervals.Count && intervals[interval].End <= offset + 0.0001D) interval++;
            if (interval < intervals.Count && offset > intervals[interval].Start + 0.0001D
                && offset < intervals[interval].End - 0.0001D) offsets.Remove(offset);
        }
    }
}
