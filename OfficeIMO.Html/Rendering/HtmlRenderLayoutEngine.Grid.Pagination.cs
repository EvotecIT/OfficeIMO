namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IEnumerable<double> FilterSafeGridBreaks(
        IEnumerable<double> sortedOffsets,
        IReadOnlyList<GridItem> items,
        IReadOnlyList<double> rowPositions,
        double contentY) {
        // Sweep item intervals alongside the sorted cuts. Only items crossing a
        // cut can reject it; unrelated rows must not multiply the operation cost.
        var intervals = items.Select(item => (
            Item: item,
            Start: contentY + rowPositions[item.Row] + item.OffsetY,
            End: contentY + rowPositions[item.Row] + item.OffsetY + item.Block!.PagedPaintExtent)).ToArray();
        var starts = intervals.OrderBy(interval => interval.Start).ToArray();
        var ends = intervals.OrderBy(interval => interval.End).ToArray();
        var active = new HashSet<GridItem>();
        var atomicVisualBottoms = new Dictionary<HtmlRenderFlowBlock, double>();
        var atomicVisualRanges = new Dictionary<HtmlRenderFlowBlock, IReadOnlyList<(double Top, double Bottom)>>();
        int startIndex = 0;
        int endIndex = 0;
        foreach (double offset in sortedOffsets) {
            CheckCancellation();
            while (startIndex < starts.Length && starts[startIndex].Start < offset - 0.0001D) {
                ChargeLayoutOperation("grid shared-row fragmentation");
                if (starts[startIndex].End > offset + 0.0001D) active.Add(starts[startIndex].Item);
                startIndex++;
            }
            while (endIndex < ends.Length && ends[endIndex].End <= offset + 0.0001D) {
                ChargeLayoutOperation("grid shared-row fragmentation");
                active.Remove(ends[endIndex].Item);
                endIndex++;
            }
            bool safe = true;
            foreach (GridItem item in active) {
                ChargeLayoutOperation("grid shared-row fragmentation");
                if (!IsSafeParallelItemBreak(item.Block!,
                        offset - contentY - rowPositions[item.Row] - item.OffsetY,
                        atomicVisualBottoms, atomicVisualRanges)) {
                    safe = false;
                    break;
                }
            }
            if (safe) yield return offset;
        }
    }
}
