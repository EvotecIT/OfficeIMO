namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Collects fragmentation facts in one source coordinate system, while the
    /// plan continues to reference the original child blocks for painting and
    /// reservation anchors. Sibling keeps can therefore move a final line rather
    /// than carrying an entire preceding paragraph into the next column.
    /// </summary>
    private ColumnFlowBreaks BuildColumnFlowBreaks(IReadOnlyList<HtmlRenderFlowBlock> children) {
        var offsets = new SortedSet<double>();
        var keeps = new List<HtmlRenderAvoidBreakRange>();
        var lines = new List<HtmlRenderLineBreakGroup>();
        var safe = new List<HtmlRenderBreakRange>();
        var forced = new List<HtmlRenderForcedBreak>();
        object[] boundaries = CreateFlowBoundaries(children);
        var boundaryPositions = boundaries.Select((boundary, index) => (boundary, index))
            .ToDictionary(item => item.boundary, item => item.index);
        var contributors = new List<(int Index, HtmlRenderForcedBreak Break)>();
        double height = 0D;
        for (int index = 0; index < children.Count; index++) {
            CheckCancellation();
            ChargeLayoutOperation("column source break facts");
            HtmlRenderFlowBlock child = children[index];
            AppendKeepWithNextRange(children, index, height, keeps);
            if (child.AvoidBreakInsideScope != HtmlFragmentationScope.None)
                keeps.Add(new HtmlRenderAvoidBreakRange(height, height + child.Height, Scope: child.AvoidBreakInsideScope));
            foreach (double offset in child.BreakOffsets) {
                ChargeLayoutOperation("column source break offset");
                if (offset > 0.0001D && offset <= child.Height + 0.0001D) offsets.Add(height + offset);
            }
            keeps.AddRange(child.AvoidBreakRanges.Select(range => range.Translate(height)));
            lines.AddRange(child.LineBreakGroups.Select(group => group.Translate(height)));
            safe.AddRange(child.SafeBreakRanges.Select(range => range.Translate(height)));
            int forcedStart = forced.Count;
            AppendChildForcedBreaks(children, index, height, forced, boundaries);
            for (int next = forcedStart; next < forced.Count; next++) contributors.Add((index, forced[next]));
            height += child.Height;
            offsets.Add(height);
        }
        HtmlRenderAvoidBreakRange[] columnKeeps = MergeColumnKeeps(keeps);
        var metrics = new HtmlRenderFlowBlock(0D, height, Array.Empty<HtmlRenderVisual>(),
            HtmlPageBreakTarget.None, HtmlPageBreakTarget.None, false, "column source breaks", offsets,
            lineBreakGroups: lines, avoidBreakRanges: columnKeeps.Concat(keeps
                .Where(range => (range.Scope & HtmlFragmentationScope.Page) != 0)
                .Select(range => range with { Scope = HtmlFragmentationScope.Page })), safeBreakRanges: safe);
        var events = new List<ColumnForcedBoundary>();
        var byBoundary = new Dictionary<object, ColumnForcedBoundary>();
        foreach ((int childIndex, HtmlRenderForcedBreak item) in contributors) {
            CheckCancellation();
            ChargeLayoutOperation("column authored break boundary");
            // Values at the column owner's first/last edge do not create an
            // extra column. A real empty sibling boundary remains Interior.
            if (item.Edge != HtmlRenderBreakEdge.Interior) continue;
            if (!byBoundary.TryGetValue(item.Boundary, out ColumnForcedBoundary? boundary)) {
                boundary = new ColumnForcedBoundary(item.Offset,
                    boundaryPositions.TryGetValue(item.Boundary, out int position) ? position : childIndex);
                byBoundary.Add(item.Boundary, boundary);
                events.Add(boundary);
            }
            if (item.Target == HtmlPageBreakTarget.Column) boundary.Column = true;
            else if (IsForcedPageBreak(item.Target)) boundary.Page = true;
        }
        return new ColumnFlowBreaks(metrics, columnKeeps, events.OrderBy(item => item.Offset).ThenBy(item => item.ChildIndex).ToArray());
    }

    private double ResolveNaturalColumnRunHeight(IReadOnlyList<HtmlRenderFlowBlock> children) {
        ColumnFlowBreaks source = BuildColumnFlowBreaks(children);
        double start = 0D, height = 0D;
        foreach (ColumnForcedBoundary boundary in source.Forced) {
            if (!boundary.Column || boundary.Page) continue;
            height = Math.Max(height, boundary.Offset - start);
            start = boundary.Offset;
        }
        return Math.Max(height, source.Metrics.Height - start);
    }

    /// <summary>Overlapping keeps form one group; touching groups can still split.</summary>
    private HtmlRenderAvoidBreakRange[] MergeColumnKeeps(IEnumerable<HtmlRenderAvoidBreakRange> ranges) {
        var merged = new List<HtmlRenderAvoidBreakRange>();
        foreach (HtmlRenderAvoidBreakRange range in ranges.Where(range =>
                     (range.Scope & HtmlFragmentationScope.Column) != 0).OrderBy(range => range.Start)) {
            CheckCancellation();
            ChargeLayoutOperation("column keep union");
            if (merged.Count > 0 && range.Start < merged[merged.Count - 1].End - 0.0001D) {
                HtmlRenderAvoidBreakRange previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = previous with {
                    End = Math.Max(previous.End, range.End), Soft = previous.Soft && range.Soft
                };
            } else merged.Add(range with { Scope = HtmlFragmentationScope.Column });
        }
        return merged.ToArray();
    }

    private static bool ColumnKeepRequiresBacktracking(ColumnFlowBreaks source, double start, double childEnd,
        double limit, double freshHeight, bool freshColumn) {
        IReadOnlyList<HtmlRenderAvoidBreakRange> ranges = source.ColumnKeeps;
        int low = 0, high = ranges.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (ranges[middle].Start < childEnd - 0.0001D) low = middle + 1;
            else high = middle;
        }
        if (low == 0) return false;
        HtmlRenderAvoidBreakRange range = ranges[low - 1];
        return start <= range.Start + 0.0001D && range.End > limit + 0.0001D
            && (range.End - range.Start <= freshHeight + 0.0001D || !freshColumn || range.Start > start + 0.0001D);
    }

    private sealed class ColumnFlowBreaks {
        internal ColumnFlowBreaks(HtmlRenderFlowBlock metrics, IReadOnlyList<HtmlRenderAvoidBreakRange> columnKeeps,
            IReadOnlyList<ColumnForcedBoundary> forced) {
            Metrics = metrics;
            ColumnKeeps = columnKeeps;
            Forced = forced;
        }
        internal HtmlRenderFlowBlock Metrics { get; }
        internal IReadOnlyList<HtmlRenderAvoidBreakRange> ColumnKeeps { get; }
        internal IReadOnlyList<ColumnForcedBoundary> Forced { get; }
    }

    private sealed class ColumnForcedBoundary {
        internal ColumnForcedBoundary(double offset, int childIndex) { Offset = offset; ChildIndex = childIndex; }
        internal int ChildIndex { get; }
        internal double Offset { get; }
        internal bool Column { get; set; }
        internal bool Page { get; set; }
    }
}
