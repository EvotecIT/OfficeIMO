using System.Globalization;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static bool TryPaintHeatmap(OfficeDrawing drawing, CartesianPanel panel, CartesianBox box,
        CartesianLayout appearance, HashSet<string> colorAxes, ref int remainingCells,
        CancellationToken cancellationToken, List<NotebookChartMark>? marks) {
        Heatmap? heat = ReadHeatmap(panel, appearance, remainingCells, colorAxes, cancellationToken);
        if (heat == null) return false;
        remainingCells -= heat.Values.Length;

        bool hasTitle = !string.IsNullOrWhiteSpace(panel.Title);
        double left = panel.Y.Title == null ? 37 : 54;
        double right = heat.Colors.Show ? 61 : 12;
        double top = hasTitle ? 26 : 14;
        double bottom = panel.X.Title == null ? 25 : 42;
        if (box.Width - left - right < 40 || box.Height - top - bottom < 36) return false;
        var plot = new CartesianBox(box.Left + left, box.Top + top, box.Width - left - right, box.Height - top - bottom);
        if (panel.X.LockAspect || panel.Y.LockAspect) {
            plot = FitHeatAspect(plot, heat.X.Maximum - heat.X.Minimum, heat.Y.Maximum - heat.Y.Minimum);
            if (plot.Width < 24 || plot.Height < 24) return false;
        }

        OfficeChartStyle style = appearance.Style;
        if (style.PlotAreaBackgroundColor is OfficeColor plotBackground)
            PaintCartesianRectangle(drawing, plot, plotBackground);
        if (hasTitle)
            PaintSceneLabel(drawing, panel.Title!, box.Left + box.Width / 2, box.Top + 10, box.Width - 12, 12, style.TitleColor);

        PaintHeatCells(drawing, plot, heat, cancellationToken, marks);
        PaintHeatAxes(drawing, box, plot, panel, heat, style);
        if (heat.Colors.Show) PaintHeatColorbar(drawing, box, plot, heat, style.TextColor);
        return true;
    }

    private static CartesianBox FitHeatAspect(CartesianBox plot, double horizontalDataSpan, double verticalDataSpan) {
        double width = Math.Min(plot.Width, plot.Height * (horizontalDataSpan / verticalDataSpan));
        double height = Math.Min(plot.Height, plot.Width * (verticalDataSpan / horizontalDataSpan));
        return new CartesianBox(plot.Left + (plot.Width - width) / 2, plot.Top + (plot.Height - height) / 2, width, height);
    }

    private static void PaintHeatCells(OfficeDrawing drawing, CartesianBox plot, Heatmap heat,
        CancellationToken cancellationToken, List<NotebookChartMark>? marks) {
        int columns = heat.X.Centers.Length;
        double[] x = HeatDrawingEdges(heat.X, plot.Left, plot.Width, vertical: false);
        double[] y = HeatDrawingEdges(heat.Y, plot.Top, plot.Height, vertical: true);
        for (int row = 0; row < heat.Y.Centers.Length; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            double top = Math.Max(plot.Top, Math.Min(y[row], y[row + 1]) + heat.VerticalGap / 2);
            double height = Math.Min(plot.Bottom, Math.Max(y[row], y[row + 1]) - heat.VerticalGap / 2) - top;
            if (height <= 0) continue;
            for (int column = 0; column < columns; column++) {
                double value = heat.Values[row * columns + column];
                if (double.IsNaN(value)) continue;
                double left = Math.Max(plot.Left, Math.Min(x[column], x[column + 1]) + heat.HorizontalGap / 2);
                double width = Math.Min(plot.Right, Math.Max(x[column], x[column + 1]) - heat.HorizontalGap / 2) - left;
                if (width <= 0) continue;
                OfficeColor color = heat.Colors.At(value);
                PaintCartesianRectangle(drawing, new CartesianBox(left, top, width, height), color);
                if (marks != null) {
                    string label = "x: " + heat.X.Label(column, precise: true) +
                        " · y: " + heat.Y.Label(row, precise: true) +
                        " · z: " + value.ToString("G", CultureInfo.InvariantCulture);
                    marks.Add(new NotebookChartMark(left, top, width, height, label));
                }
                if (heat.LabelFormat != null && width > 19 && height > 12) {
                    string label = value.ToString(heat.LabelFormat, CultureInfo.InvariantCulture);
                    int luminance = 299 * color.R + 587 * color.G + 114 * color.B;
                    OfficeColor textColor = luminance > 141000 ? OfficeColor.Black : OfficeColor.White;
                    PaintSceneLabel(drawing, label, left + width / 2, top + height / 2, width - 3, 10, textColor);
                }
            }
        }
    }

    private static double[] HeatDrawingEdges(HeatAxis axis, double origin, double extent, bool vertical) {
        var result = new double[axis.Edges.Length];
        for (int index = 0; index < result.Length; index++) {
            double normalized = axis.Fraction(axis.Edges[index]);
            result[index] = origin + (vertical ? 1 - normalized : normalized) * extent;
        }
        return result;
    }

    private static void PaintHeatAxes(OfficeDrawing drawing, CartesianBox box, CartesianBox plot, CartesianPanel panel,
        Heatmap heat, OfficeChartStyle style) {
        OfficeColor axis = style.AxisColor;
        OfficeColor text = style.TextColor;
        PaintSceneLine(drawing, new OfficePoint(plot.Left, plot.Bottom), new OfficePoint(plot.Right, plot.Bottom), axis, 0.6);
        PaintSceneLine(drawing, new OfficePoint(plot.Left, plot.Top), new OfficePoint(plot.Left, plot.Bottom), axis, 0.6);
        if (panel.X.ShowLabels) {
            int stride = HeatTickStride(heat.X.Centers.Length, plot.Width, minStep: 44);
            for (int column = 0; column < heat.X.Centers.Length; column += stride) {
                double t = heat.X.Fraction(heat.X.Centers[column]);
                if (t < 0 || t > 1) continue;
                double x = plot.Left + t * plot.Width;
                PaintSceneLine(drawing, new OfficePoint(x, plot.Bottom), new OfficePoint(x, plot.Bottom + 3), axis, 0.5);
                PaintSceneLabel(drawing, heat.X.Label(column), x, plot.Bottom + 10, 42, 8, text);
            }
        }
        if (panel.Y.ShowLabels) {
            int stride = HeatTickStride(heat.Y.Centers.Length, plot.Height, minStep: 15);
            for (int row = 0; row < heat.Y.Centers.Length; row += stride) {
                double t = heat.Y.Fraction(heat.Y.Centers[row]);
                if (t < 0 || t > 1) continue;
                double y = plot.Bottom - t * plot.Height;
                PaintSceneLine(drawing, new OfficePoint(plot.Left - 3, y), new OfficePoint(plot.Left, y), axis, 0.5);
                drawing.AddText(heat.Y.Label(row), plot.Left - 35, y - 5, 29, 10, new OfficeFontInfo(null, 8), text,
                    OfficeTextAlignment.Right, shrinkToFit: true);
            }
        }
        if (!string.IsNullOrWhiteSpace(panel.X.Title)) {
            PaintSceneLabel(drawing, panel.X.Title!, plot.Left + plot.Width / 2, plot.Bottom + 26, plot.Width, 10, text);
        }
        if (!string.IsNullOrWhiteSpace(panel.Y.Title)) {
            double labelHeight = Math.Min(plot.Height, 110);
            double pivotX = box.Left + 8;
            double pivotY = plot.Top + plot.Height / 2;
            var title = new OfficeDrawing(labelHeight, 14);
            title.AddText(panel.Y.Title!, 0, 0, labelHeight, 14, new OfficeFontInfo(null, 10), text,
                OfficeTextAlignment.Center, shrinkToFit: true);
            drawing.AddEffectDrawing(title, OfficeTransform.Translate(-labelHeight / 2, -7)
                .Then(OfficeTransform.RotateDegrees(-90)).Then(OfficeTransform.Translate(pivotX, pivotY)));
        }
    }

    private static int HeatTickStride(int count, double extent, double minStep) =>
        Math.Max(1, (int)Math.Ceiling(count * minStep / extent));

    private static void PaintHeatColorbar(OfficeDrawing drawing, CartesianBox box, CartesianBox plot, Heatmap heat,
        OfficeColor textColor) {
        const int steps = 24;
        double left = plot.Right + 12;
        double top = plot.Top + 11;
        double height = plot.Height - 15;
        double step = height / steps;
        for (int index = 0; index < steps; index++) {
            double fraction = 1 - (index + 0.5) / steps;
            PaintCartesianRectangle(drawing, new CartesianBox(left, top + index * step, 9, step + 0.03),
                heat.Colors.FromFraction(fraction));
        }
        PaintHeatColorbarLabel(drawing, FormatSceneNumber(heat.Colors.Max), left + 12, top, textColor);
        if (heat.Colors.Min != heat.Colors.Max) {
            PaintHeatColorbarLabel(drawing, FormatSceneNumber((heat.Colors.Min + heat.Colors.Max) / 2), left + 12, top + height / 2, textColor);
            PaintHeatColorbarLabel(drawing, FormatSceneNumber(heat.Colors.Min), left + 12, top + height, textColor);
        }
        if (!string.IsNullOrWhiteSpace(heat.ColorbarTitle)) {
            double labelLeft = Math.Min(left - 6, box.Right - 56);
            drawing.AddText(heat.ColorbarTitle!, labelLeft, top - 13, 56, 12, new OfficeFontInfo(null, 8), textColor,
                OfficeTextAlignment.Left, shrinkToFit: true);
        }
    }

    private static void PaintHeatColorbarLabel(OfficeDrawing drawing, string label, double left, double top, OfficeColor color) =>
        drawing.AddText(label, left, top - 5, 35, 10, new OfficeFontInfo(null, 8), color, shrinkToFit: true);
}
