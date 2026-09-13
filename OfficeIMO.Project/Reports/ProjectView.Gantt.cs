using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private void DrawGanttLinks(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount,
        double chartX, double top, double cellWidth, double[] heights) {
        var tops = new double[heights.Length]; double current = top;
        for (int i = 0; i < heights.Length; i++) { tops[i] = current; current += heights[i]; }
        var locations = Rows.Skip(rowOffset).Take(rowCount).Select((row, index) => (row, index)).ToDictionary(item => item.row.Uid);
        DateTime first = Buckets[bucketOffset].Start, last = Buckets[bucketOffset + bucketCount - 1].Finish;
        double right = chartX + bucketCount * cellWidth;
        foreach (var link in Links) {
            if (!locations.TryGetValue(link.PredecessorUid, out var source) || !locations.TryGetValue(link.SuccessorUid, out var target)) continue;
            if (!HasVisibleBar(source.row, first, last) || !HasVisibleBar(target.row, first, last)) continue;
            bool fromFinish = link.Type == ProjectDependencyType.FinishToStart || link.Type == ProjectDependencyType.FinishToFinish;
            bool toFinish = link.Type == ProjectDependencyType.FinishToFinish || link.Type == ProjectDependencyType.StartToFinish;
            DateTime? from = fromFinish ? source.row.Finish : source.row.Start, to = toFinish ? target.row.Finish : target.row.Start;
            if (!from.HasValue || !to.HasValue || from < first || from > last || to < first || to > last) continue;
            double sx = chartX + TimePosition(from.Value, bucketOffset, bucketCount, cellWidth);
            double tx = chartX + TimePosition(to.Value, bucketOffset, bucketCount, cellWidth);
            if (source.row.Start == source.row.Finish) sx = MilestoneAnchor(sx, chartX, right, source.row.IsSummary ? 8 : 13, fromFinish);
            if (target.row.Start == target.row.Finish) tx = MilestoneAnchor(tx, chartX, right, target.row.IsSummary ? 8 : 13, toFinish);
            double sy = tops[source.index] + heights[source.index] / 2 - 3, ty = tops[target.index] + heights[target.index] / 2 - 3;
            double sourceGutter = Math.Max(chartX, Math.Min(right, sx + (fromFinish ? 5 : -5)));
            double targetGutter = Math.Max(chartX, Math.Min(right, tx + (toFinish ? 5 : -5)));
            double routeY = tops[source.index] + (target.index > source.index ? heights[source.index] - 3 : 2);
            Line(drawing, sx, sy, sourceGutter, sy, false);
            Line(drawing, sourceGutter, sy, sourceGutter, routeY, false);
            Line(drawing, sourceGutter, routeY, targetGutter, routeY, false);
            Line(drawing, targetGutter, routeY, targetGutter, ty, false);
            Line(drawing, targetGutter, ty, tx, ty, true);
        }
    }

    private static double MilestoneAnchor(double position, double left, double right, double size, bool finish) {
        double center = Math.Max(left + size / 2, Math.Min(right - size / 2, position));
        return center + (finish ? size / 2 : -size / 2);
    }

    private static bool HasVisibleBar(ProjectViewRow row, DateTime first, DateTime last) =>
        row.Start.HasValue && row.Finish.HasValue && row.Finish >= row.Start && row.Start < last &&
        (row.Finish > first || row.Start == row.Finish && row.Finish == first);
}
