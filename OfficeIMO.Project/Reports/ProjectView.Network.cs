using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private IReadOnlyList<ProjectViewPage> RenderNetwork(CancellationToken token) {
        double width = Layout.PageWidth - 2 * Layout.Margin;
        int columns = Math.Max(1, (int)(width / 240));
        int lines = Math.Max(1, (int)((Layout.PageHeight - 2 * Layout.Margin - 100) / 100));
        int perPage = columns * lines;
        int pageCount = Math.Max(1, (Rows.Count + perPage - 1) / perPage);
        if (pageCount > Layout.MaxPages) throw new InvalidOperationException("Report exceeds MaxPages.");
        var pages = new List<ProjectViewPage>();
        var locations = Rows.Select((r, i) => (r.Uid, Index: i)).ToDictionary(r => r.Uid, r => r.Index);
        var incoming = Links.ToLookup(l => l.SuccessorUid);
        for (int p = 0; p < pageCount; p++) {
            token.ThrowIfCancellationRequested();
            var drawing = NewPage(p + 1, pageCount);
            int offset = p * perPage, count = Math.Min(perPage, Rows.Count - offset);
            double cell = width / columns;
            // Route local links through the gutters before painting nodes. Cross-page links retain explicit page references below.
            foreach (var link in Links) {
                int source = locations[link.PredecessorUid] - offset, target = locations[link.SuccessorUid] - offset;
                if (source < 0 || source >= count || target < 0 || target >= count) continue;
                token.ThrowIfCancellationRequested();
                double sx = Layout.Margin + source % columns * cell + cell - 16, sy = Layout.Margin + 101 + source / columns * 100;
                double tx = Layout.Margin + target % columns * cell, ty = Layout.Margin + 101 + target / columns * 100;
                if (source / columns == target / columns && target % columns == source % columns + 1) Line(drawing, sx, sy, tx, ty, true);
                else {
                    double gutter = sx + 8, destinationGutter = tx - 8;
                    double routeY = Layout.Margin + 60 + source / columns * 100 + (target / columns > source / columns ? 90 : -8);
                    Line(drawing, sx, sy, gutter, sy, false);
                    Line(drawing, gutter, sy, gutter, routeY, false);
                    Line(drawing, gutter, routeY, destinationGutter, routeY, false);
                    Line(drawing, destinationGutter, routeY, destinationGutter, ty, false);
                    Line(drawing, destinationGutter, ty, tx, ty, true);
                }
            }
            for (int i = 0; i < count; i++) {
                var row = Rows[offset + i];
                double x = Layout.Margin + i % columns * cell, y = Layout.Margin + 60 + i / columns * 100;
                Rect(drawing, x, y, cell - 16, 82, OfficeColor.ParseHex("#EFF4F8"));
                Rect(drawing, x, y, 4, 82, row.IsSummary ? Ink : row.IsCritical ? Critical : Accent);
                Text(drawing, row.Uid + "  " + row.Name, x + 10, y + 7, cell - 36, 20, 11, Ink, true);
                Text(drawing, row.Start?.ToString("yyyy-MM-dd") + " → " + row.Finish?.ToString("yyyy-MM-dd"), x + 10, y + 30, cell - 36, 17, 9, Muted);
                var predecessors = incoming[row.Uid].Select(l => l.PredecessorUid + " " + l.Type + (l.Lag.HasValue || l.LagPercent.HasValue ? " " + l.LagText : "") + " (p" + (locations[l.PredecessorUid] / perPage + 1) + ")");
                Text(drawing, "From: " + string.Join(", ", predecessors.DefaultIfEmpty("none")), x + 10, y + 51, cell - 36, 26, 8, Ink);
            }
            if (count == 0) Text(drawing, "No rows match this selection.", Layout.Margin, Layout.Margin + 70, width, 25, 12, Muted);
            pages.Add(new ProjectViewPage(drawing, offset, count, 0, 0));
        }
        return pages.AsReadOnly();
    }

    private static void Line(OfficeDrawing drawing, double x1, double y1, double x2, double y2, bool arrow) {
        if (x1 == x2 && y1 == y2) return;
        var shape = OfficeShape.Line(x1, y1, x2, y2); shape.StrokeColor = Muted; shape.StrokeWidth = 1;
        if (arrow) shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 5, 5);
        drawing.AddShape(shape, Math.Min(x1, x2), Math.Min(y1, y2));
    }
}
