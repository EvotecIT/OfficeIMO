using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private sealed class NetworkPage {
        internal int[] Indices = Array.Empty<int>();
        internal readonly Dictionary<int, (double X, double Y, double Height)> Nodes = new();
        internal double Height;
    }

    private IReadOnlyList<ProjectViewPage> RenderDependencyNetwork(CancellationToken token) {
        double width = PageWidth - 2 * PageMargin;
        const double gap = 42, padding = 14;
        bool compact = BodyHeight < 190;
        double titleTop = compact ? 27 : 35, referenceOffset = compact ? 28 : 39;
        double progressOffset = compact ? 20 : 25, verticalGap = compact ? 12 : 24;
        int columns = Math.Max(1, (int)((width + gap) / (230 + gap)));
        double nodeWidth = (width - (columns - 1) * gap) / columns;
        var stages = NetworkStages(token);
        var locations = Rows.Select((row, index) => (row.Uid, index)).ToDictionary(item => item.Uid, item => item.index);
        var incoming = Links.ToLookup(link => link.SuccessorUid);
        var outgoing = Links.ToLookup(link => link.PredecessorUid);
        var pageNumbers = new int[Rows.Count];
        string LinkLabel(ProjectViewLink link, bool from, bool measure) {
            int index = locations[from ? link.PredecessorUid : link.SuccessorUid];
            string number = measure ? new string('8', MaxPages.ToString(CultureInfo.InvariantCulture).Length) : pageNumbers[index].ToString(CultureInfo.InvariantCulture);
            return "#" + Rows[index].Uid + " " + DependencyCode(link.Type) + (link.Lag.HasValue || link.LagPercent.HasValue ? " " + link.LagText : "") + " · p" + number;
        }
        string References(int index, bool measure) {
            string from = string.Join(", ", incoming[Rows[index].Uid].Select(link => LinkLabel(link, true, measure)));
            string to = string.Join(", ", outgoing[Rows[index].Uid].Select(link => LinkLabel(link, false, measure)));
            return (from.Length == 0 ? "Start of sequence" : "From " + from) + (to.Length == 0 ? "\nEnd of sequence" : "\nTo " + to);
        }
        var titleHeights = Rows.Select(row => MeasureTextHeight(row.Name, 13, nodeWidth - 2 * padding, true)).ToArray();
        var referenceHeights = Enumerable.Range(0, Rows.Count).Select(index => MeasureTextHeight(References(index, true), 9, nodeWidth - 2 * padding)).ToArray();
        var nodeHeights = Enumerable.Range(0, Rows.Count).Select(index => titleHeights[index] + referenceHeights[index] + (compact ? 75 : 104)).ToArray();
        var plans = new List<NetworkPage>();
        for (int firstStage = 0; firstStage < stages.Length; firstStage += columns) {
            token.ThrowIfCancellationRequested();
            var panel = stages.Skip(firstStage).Take(columns).ToArray();
            int lanes = panel.Max(stage => stage.Length);
            var slotHeights = new double[lanes];
            var slots = new Dictionary<int, (int Column, int Lane)>();
            for (int column = 0; column < panel.Length; column++) {
                int firstLane = (lanes - panel[column].Length) / 2;
                for (int lane = 0; lane < panel[column].Length; lane++) {
                    int index = panel[column][lane], slot = firstLane + lane;
                    slots.Add(index, (column, slot));
                    slotHeights[slot] = Math.Max(slotHeights[slot], nodeHeights[index] + verticalGap);
                }
            }
            var byLane = slots.ToLookup(item => item.Value.Lane);
            foreach (var slice in OfficeTablePagination.Paginate(slotHeights, BodyHeight, 0, MaxPages, token)) {
                if (plans.Count >= MaxPages) throw new InvalidOperationException("Report exceeds MaxPages.");
                var plan = new NetworkPage { Height = slice.Height - verticalGap };
                double top = 0;
                for (int lane = slice.RowOffset; lane < slice.RowOffset + slice.RowCount; lane++) {
                    token.ThrowIfCancellationRequested();
                    foreach (var item in byLane[lane].OrderBy(item => item.Value.Column)) {
                        int index = item.Key;
                        plan.Nodes.Add(index, (item.Value.Column * (nodeWidth + gap), top, nodeHeights[index]));
                        pageNumbers[index] = plans.Count + 1;
                    }
                    top += slotHeights[lane];
                }
                plan.Indices = plan.Nodes.Keys.ToArray();
                plans.Add(plan);
            }
        }
        if (plans.Count == 0) plans.Add(new NetworkPage { Height = 36 });
        var result = new List<ProjectViewPage>();
        for (int p = 0; p < plans.Count; p++) {
            token.ThrowIfCancellationRequested();
            var plan = plans[p]; var drawing = NewPage(p + 1, plans.Count, plan.Height);
            double originY = PageMargin + TitleAreaHeight;
            foreach (var link in Links) {
                token.ThrowIfCancellationRequested();
                if (!plan.Nodes.TryGetValue(locations[link.PredecessorUid], out var from) || !plan.Nodes.TryGetValue(locations[link.SuccessorUid], out var to)) continue;
                double sx = PageMargin + from.X + nodeWidth, sy = originY + from.Y + from.Height / 2;
                double tx = PageMargin + to.X, ty = originY + to.Y + to.Height / 2;
                double gutter = sx + gap / 2;
                var color = Rows[locations[link.PredecessorUid]].IsCritical && Rows[locations[link.SuccessorUid]].IsCritical ? Critical : Muted;
                NetworkLine(drawing, sx, sy, gutter, sy, color);
                if (tx - sx > gap + 1) {
                    double routeY = originY - 10, targetGutter = tx - gap / 2;
                    NetworkLine(drawing, gutter, sy, gutter, routeY, color);
                    NetworkLine(drawing, gutter, routeY, targetGutter, routeY, color);
                    NetworkLine(drawing, targetGutter, routeY, targetGutter, ty, color);
                    NetworkLine(drawing, targetGutter, ty, tx, ty, color, true);
                } else {
                    NetworkLine(drawing, gutter, sy, gutter, ty, color);
                    NetworkLine(drawing, gutter, ty, tx, ty, color, true);
                }
            }
            foreach (int index in plan.Indices) {
                token.ThrowIfCancellationRequested();
                var row = Rows[index]; var node = plan.Nodes[index];
                double x = PageMargin + node.X, y = originY + node.Y;
                var color = row.IsCritical ? Critical : row.IsSummary ? Ink : Accent;
                var card = OfficeShape.RoundedRectangle(nodeWidth, node.Height, 6);
                card.FillColor = OfficeColor.White; card.StrokeColor = Grid; card.StrokeWidth = 1;
                drawing.AddShape(card, x, y);
                Rect(drawing, x + padding, y + 13, 3, 12, color);
                string status = row.PercentComplete >= 100 ? "COMPLETE" : row.IsSummary ? "SUMMARY" : row.IsCritical ? "CRITICAL" : "TASK";
                Text(drawing, "#" + row.Uid + "  ·  " + status, x + padding + 10, y + 12, nodeWidth - 2 * padding - 10, 14, 9, color, true);
                Text(drawing, row.Name, x + padding, y + titleTop, nodeWidth - 2 * padding, titleHeights[index], 13, Ink, true, true);
                double dateY = y + titleTop + 8 + titleHeights[index];
                string dates = (row.Start?.ToString("dd MMM yy", CultureInfo.InvariantCulture) ?? "—") + "  →  " + (row.Finish?.ToString("dd MMM yy", CultureInfo.InvariantCulture) ?? "—");
                Text(drawing, dates, x + padding, dateY, nodeWidth - 2 * padding, 16, 10, Muted);
                Rect(drawing, x + padding, dateY + progressOffset, nodeWidth - 2 * padding, 3, Grid);
                if (Layout.ShowProgress && row.PercentComplete.HasValue) {
                    double progress = Math.Max(0, Math.Min(100, row.PercentComplete.Value)) / 100d;
                    if (progress > 0) Rect(drawing, x + padding, dateY + progressOffset, (nodeWidth - 2 * padding) * progress, 3, color);
                }
                Text(drawing, References(index, false), x + padding, dateY + referenceOffset, nodeWidth - 2 * padding, referenceHeights[index], 9, Muted, wrap: true);
            }
            if (plan.Indices.Length == 0) Text(drawing, "No rows match this selection.", PageMargin, originY + 8, width, 24, 12, Muted);
            result.Add(new ProjectViewPage(drawing, plan.Indices.Length == 0 ? 0 : plan.Indices.Min(), plan.Indices.Length, 0, 0, plan.Indices));
        }
        return result.AsReadOnly();
    }

    private static void NetworkLine(OfficeDrawing drawing, double x1, double y1, double x2, double y2, OfficeColor color, bool arrow = false) {
        if (x1 == x2 && y1 == y2) return;
        var line = OfficeShape.Line(x1, y1, x2, y2); line.StrokeColor = color; line.StrokeWidth = 1.25;
        if (arrow) line.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 6, 6);
        drawing.AddShape(line, Math.Min(x1, x2), Math.Min(y1, y2));
    }
}
