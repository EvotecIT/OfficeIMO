using OfficeIMO.Drawing;

namespace OfficeIMO.Project;

public sealed partial class ProjectView {
    private IReadOnlyList<ProjectViewPage> RenderNetwork(CancellationToken token) {
        return Layout.NetworkLayout == ProjectNetworkLayout.Compact ? RenderCompactNetwork(token) : RenderDependencyNetwork(token);
    }

    private IReadOnlyList<ProjectViewPage> RenderCompactNetwork(CancellationToken token) {
        double width = Layout.PageWidth - 2 * Layout.Margin;
        int columns = Math.Max(1, (int)(width / 240));
        double cell = width / columns;
        var pages = new List<ProjectViewPage>();
        var locations = Rows.Select((r, i) => (r.Uid, Index: i)).ToDictionary(r => r.Uid, r => r.Index);
        var incoming = Links.ToLookup(l => l.SuccessorUid);
        var titleHeights = new double[Rows.Count]; var dateHeights = new double[Rows.Count];
        var rowHeights = new double[(Rows.Count + columns - 1) / columns];
        var pageNumbers = new int[Rows.Count];
        string Predecessors(int index, bool measuring) => "From: " + string.Join(", ", incoming[Rows[index].Uid].Select(link =>
            link.PredecessorUid + " " + DependencyCode(link.Type) + (link.Lag.HasValue || link.LagPercent.HasValue ? " " + link.LagText : "")
            + " (p" + (measuring ? new string('8', Layout.MaxPages.ToString(System.Globalization.CultureInfo.InvariantCulture).Length) : pageNumbers[locations[link.PredecessorUid]].ToString(System.Globalization.CultureInfo.InvariantCulture)) + ")").DefaultIfEmpty("none"));
        string Dates(int index) => Rows[index].Start?.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture) + " → "
            + Rows[index].Finish?.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture);
        for (int i = 0; i < Rows.Count; i++) {
            token.ThrowIfCancellationRequested();
            titleHeights[i] = MeasureTextHeight(Rows[i].Uid + "  " + Rows[i].Name, 11, cell - 36, true);
            dateHeights[i] = MeasureTextHeight(Dates(i), 9, cell - 36);
            double height = titleHeights[i] + dateHeights[i] + MeasureTextHeight(Predecessors(i, true), 9, cell - 36) + 52;
            rowHeights[i / columns] = Math.Max(rowHeights[i / columns], height);
        }
        var rowPages = OfficeTablePagination.Paginate(rowHeights, BodyHeight, 0, Layout.MaxPages, token);
        int pageCount = rowPages.Count;
        for (int p = 0; p < pageCount; p++)
            for (int i = rowPages[p].RowOffset * columns; i < Math.Min(Rows.Count, (rowPages[p].RowOffset + rowPages[p].RowCount) * columns); i++) pageNumbers[i] = p + 1;
        for (int p = 0; p < pageCount; p++) {
            token.ThrowIfCancellationRequested();
            var drawing = NewPage(p + 1, pageCount, Math.Max(36, rowPages[p].Height));
            int rowOffset = rowPages[p].RowOffset;
            int offset = rowOffset * columns, count = Math.Min(rowPages[p].RowCount * columns, Rows.Count - offset);
            var tops = new double[rowPages[p].RowCount]; double top = Layout.Margin + TitleAreaHeight;
            for (int r = 0; r < tops.Length; r++) { tops[r] = top; top += rowHeights[rowOffset + r]; }
            // Route local links through the gutters before painting nodes. Cross-page links retain explicit page references below.
            foreach (var link in Links) {
                int source = locations[link.PredecessorUid] - offset, target = locations[link.SuccessorUid] - offset;
                if (source < 0 || source >= count || target < 0 || target >= count) continue;
                token.ThrowIfCancellationRequested();
                double sx = Layout.Margin + source % columns * cell + cell - 16, sy = tops[source / columns] + (rowHeights[rowOffset + source / columns] - 20) / 2;
                double tx = Layout.Margin + target % columns * cell, ty = tops[target / columns] + (rowHeights[rowOffset + target / columns] - 20) / 2;
                if (source / columns == target / columns && target % columns == source % columns + 1) Line(drawing, sx, sy, tx, ty, true);
                else {
                    double gutter = sx + 8, destinationGutter = tx - 8;
                    double routeY = tops[source / columns] + (target / columns > source / columns ? rowHeights[rowOffset + source / columns] - 10 : -8);
                    Line(drawing, sx, sy, gutter, sy, false);
                    Line(drawing, gutter, sy, gutter, routeY, false);
                    Line(drawing, gutter, routeY, destinationGutter, routeY, false);
                    Line(drawing, destinationGutter, routeY, destinationGutter, ty, false);
                    Line(drawing, destinationGutter, ty, tx, ty, true);
                }
            }
            for (int i = 0; i < count; i++) {
                var row = Rows[offset + i];
                double x = Layout.Margin + i % columns * cell, y = tops[i / columns], height = rowHeights[rowOffset + i / columns] - 20;
                Rect(drawing, x, y, cell - 16, height, OfficeColor.ParseHex("#EFF4F8"));
                Rect(drawing, x, y, 4, height, row.IsSummary ? Ink : row.IsCritical ? Critical : Accent);
                Text(drawing, row.Uid + "  " + row.Name, x + 10, y + 10, cell - 36, titleHeights[offset + i], 11, Ink, true, true);
                double dateY = y + 16 + titleHeights[offset + i];
                Text(drawing, Dates(offset + i), x + 10, dateY, cell - 36, dateHeights[offset + i], 9, Muted, wrap: true);
                double linksY = dateY + dateHeights[offset + i] + 6;
                Text(drawing, Predecessors(offset + i, false), x + 10, linksY, cell - 36, y + height - linksY - 10, 9, Ink, wrap: true);
            }
            if (count == 0) Text(drawing, "No rows match this selection.", Layout.Margin, Layout.Margin + TitleAreaHeight + 10, width, 25, 12, Muted);
            pages.Add(new ProjectViewPage(drawing, offset, count, 0, 0));
        }
        return pages.AsReadOnly();
    }

    private static string DependencyCode(ProjectDependencyType type) => type switch {
        ProjectDependencyType.FinishToStart => "FS", ProjectDependencyType.StartToStart => "SS",
        ProjectDependencyType.FinishToFinish => "FF", ProjectDependencyType.StartToFinish => "SF", _ => type.ToString()
    };

    private static void Line(OfficeDrawing drawing, double x1, double y1, double x2, double y2, bool arrow) {
        if (x1 == x2 && y1 == y2) return;
        var shape = OfficeShape.Line(x1, y1, x2, y2); shape.StrokeColor = Muted; shape.StrokeWidth = 1;
        if (arrow) shape.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 5, 5);
        drawing.AddShape(shape, Math.Min(x1, x2), Math.Min(y1, y2));
    }
}
