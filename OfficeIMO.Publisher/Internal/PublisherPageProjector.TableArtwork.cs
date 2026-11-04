using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherPageProjector {
    /// <summary>Native utility edges share grid boundaries. Outer strokes stay inside the table; horizontal edges paint over vertical edges.</summary>
    private static void ProjectTableArtwork(PublisherTable table, OfficeDrawing drawing) {
        foreach (PublisherTableCell cell in table.Cells) {
            if (!cell.FillColor.HasValue) continue;
            OfficeShape fill = OfficeShape.Rectangle(cell.Width, cell.Height);
            fill.FillColor = cell.FillColor; fill.StrokeWidth = 0;
            drawing.AddShape(fill, cell.X, cell.Y);
        }
        double[] xs = Boundaries(table.ColumnWidths), ys = Boundaries(table.RowHeights);
        foreach (PublisherTableBorderSegment border in table.BorderSegments.OrderBy(item => item.FirstRow == item.LastRow)) {
            bool horizontal = border.FirstRow == border.LastRow;
            double x1 = xs[border.FirstColumn], x2 = xs[border.LastColumn];
            double y1 = ys[border.FirstRow], y2 = ys[border.LastRow];
            double half = border.Side.Width / 2;
            if (horizontal) {
                if (border.FirstRow == 0) y1 = y2 = half;
                else if (border.FirstRow == table.RowCount) y1 = y2 = ys[ys.Length - 1] - half;
            } else {
                if (border.FirstColumn == 0) x1 = x2 = half;
                else if (border.FirstColumn == table.ColumnCount) x1 = x2 = xs[xs.Length - 1] - half;
            }
            OfficeShape line = OfficeShape.Line(x1, y1, x2, y2);
            line.StrokeColor = border.Side.Color; line.StrokeWidth = border.Side.Width;
            drawing.AddShape(line, Math.Min(x1, x2), Math.Min(y1, y2));
        }

        static double[] Boundaries(IReadOnlyList<double> tracks) {
            var result = new double[tracks.Count + 1];
            for (int i = 0; i < tracks.Count; i++) result[i + 1] = result[i] + tracks[i];
            return result;
        }
    }
}
