using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private void DrawGanttLinks(OfficeDrawing drawing, int rowOffset, int rowCount, int bucketOffset, int bucketCount,
        double chartX, double top, double cellWidth) {
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
            double sy = top + source.index * 30 + (source.row.IsSummary ? 11 : 13), ty = top + target.index * 30 + (target.row.IsSummary ? 11 : 13);
            double sourceGutter = Math.Max(chartX, Math.Min(right, sx + (fromFinish ? 5 : -5)));
            double targetGutter = Math.Max(chartX, Math.Min(right, tx + (toFinish ? 5 : -5)));
            double routeY = top + source.index * 30 + (target.index > source.index ? 28 : 2);
            Line(drawing, sx, sy, sourceGutter, sy, false);
            Line(drawing, sourceGutter, sy, sourceGutter, routeY, false);
            Line(drawing, sourceGutter, routeY, targetGutter, routeY, false);
            Line(drawing, targetGutter, routeY, targetGutter, ty, false);
            Line(drawing, targetGutter, ty, tx, ty, true);
        }
    }

    private static bool HasVisibleBar(ProjectViewRow row, DateTime first, DateTime last) =>
        row.Start.HasValue && row.Finish.HasValue && row.Finish >= row.Start && row.Start < last &&
        (row.Finish > first || row.Start == row.Finish && row.Finish == first);
}
