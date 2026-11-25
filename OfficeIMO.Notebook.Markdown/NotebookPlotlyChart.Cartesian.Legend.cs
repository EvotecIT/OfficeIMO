using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static CartesianBox? ReserveCartesianLegend(int itemCount, OfficeChartLegendPosition position,
        ref CartesianBox content) {
        if (itemCount == 0) return null;
        if (position == OfficeChartLegendPosition.Bottom) {
            int columns = CartesianLegendColumns(content.Width, itemCount);
            int rows = (itemCount + columns - 1) / columns;
            double height = 14 * rows + 3;
            var band = new CartesianBox(content.Left, content.Bottom - height, content.Width, height);
            content = content with { Height = content.Height - height };
            return band;
        }
        double width = Math.Min(85, content.Width * 0.19);
        var column = new CartesianBox(content.Right - width, content.Top + 18, width, content.Height - 18);
        content = content with { Width = content.Width - width };
        return column;
    }

    private static void PaintCartesianLegend(OfficeDrawing drawing, CartesianBox box, OfficeChartLegendPosition position,
        List<CartesianLegendItem> entries, OfficeColor textColor) {
        bool horizontal = position == OfficeChartLegendPosition.Bottom;
        int columns = horizontal ? CartesianLegendColumns(box.Width, entries.Count) : 1;
        int rows = (entries.Count + columns - 1) / columns;
        double itemWidth = box.Width / columns;
        double itemHeight = Math.Min(14, box.Height / rows);
        entries.Sort((left, right) => left.TraceIndex.CompareTo(right.TraceIndex));
        for (int i = 0; i < entries.Count; i++) {
            CartesianLegendItem entry = entries[i];
            double x = box.Left + i % columns * itemWidth + 4;
            double y = box.Top + (i / columns + 0.5) * itemHeight;
            if (entry.Bar) {
                PaintCartesianRectangle(drawing, new CartesianBox(x + 1, y - 3, 7, 6), entry.Color);
            } else {
                if (entry.Line) PaintSceneLine(drawing, new OfficePoint(x, y), new OfficePoint(x + 10, y), entry.LineColor, 1.25);
                if (entry.Markers) PaintSceneMarker(drawing, new OfficePoint(x + 5, y), entry.Color, 4,
                    entry.MarkerShape ?? OfficeChartMarkerShape.Circle);
            }
            drawing.AddText(entry.Label, x + 14, y - itemHeight / 2, itemWidth - 20, itemHeight,
                new OfficeFontInfo(null, 8), textColor, verticalAlignment: OfficeTextVerticalAlignment.Center, shrinkToFit: true);
        }
    }

    private static int CartesianLegendColumns(double width, int itemCount) =>
        Math.Min(itemCount, Math.Max(1, (int)(width / 100)));

    private readonly record struct CartesianLegendItem(int TraceIndex, string Label, OfficeColor Color, OfficeColor LineColor,
        bool Bar, bool Line, bool Markers, OfficeChartMarkerShape? MarkerShape);
}
