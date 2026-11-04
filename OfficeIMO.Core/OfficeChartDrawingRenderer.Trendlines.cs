using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Drawing;

public static partial class OfficeChartDrawingRenderer {
    private const int TrendlineSamples = 64;

    // Draws series trendlines over the plot. Category charts fit against the 1-based category
    // number, as Office does; scatter charts fit against the x values.
    private static void AddTrendlines(OfficeDrawing drawing, OfficeChartSnapshot snapshot, ValueRange primaryRange,
        ValueRange secondaryRange, bool hasSecondaryAxis, double plotLeft, double plotTop, double plotWidth,
        double plotHeight, OfficeChartStyle style, OfficeChartLayout layout) {
        if (IsBarChart(snapshot.ChartKind) || !snapshot.Data.Series.Any(series => series.Trendline != null)) return;
        IReadOnlyList<string> categories = snapshot.Data.Categories;
        bool scatter = IsScatterChart(snapshot.ChartKind);
        IReadOnlyList<double> sharedX = scatter ? GetScatterXValues(categories) : Array.Empty<double>();
        ValueRange xRange = default;
        if (scatter) {
            List<OfficeSeries> scatterSeries = GetRenderableScatterSeries(snapshot).Select(item => item.Series).ToList();
            xRange = ApplyValueAxisScale(GetScatterPointRanges(scatterSeries, sharedX).XRange, layout, horizontal: true);
        }

        var geometry = new OfficeDrawing(drawing.Width, drawing.Height);
        var labels = new List<(string Text, double X, double Y)>();
        for (int seriesIndex = 0; seriesIndex < snapshot.Data.Series.Count; seriesIndex++) {
            OfficeSeries series = snapshot.Data.Series[seriesIndex];
            OfficeChartTrendline? trendline = series.Trendline;
            if (trendline == null) continue;
            OfficeChartType kind = GetEffectiveSeriesKind(snapshot, series);
            if (IsStackedBarOrColumnChart(kind) || IsPercentStackedBarOrColumnChart(kind) || IsStackedLineChart(kind) ||
                IsPercentStackedLineChart(kind) || IsStackedAreaChart(kind) || IsPercentStackedAreaChart(kind)) continue;
            ValueRange range = hasSecondaryAxis && series.AxisGroup == OfficeChartAxisGroup.Secondary ? secondaryRange : primaryRange;

            var xs = new List<double>();
            var ys = new List<double>();
            IReadOnlyList<double>? seriesX = scatter ? series.XValues ?? sharedX : null;
            for (int i = 0; i < series.Values.Count; i++) {
                if (!TryGetSeriesValue(series, i, out double y)) continue;
                double x = seriesX != null ? (i < seriesX.Count ? seriesX[i] : double.NaN) : i + 1D;
                if (!IsFiniteChartValue(x)) continue;
                xs.Add(x);
                ys.Add(y);
            }

            if (xs.Count < 2) continue;
            Func<double, double> toPixelX = scatter
                ? value => plotLeft + plotWidth * GetUnclampedPlotRatio(value, xRange.Min, xRange.Max)
                : value => CategoryNumberToPixel(snapshot, plotLeft, plotWidth, value, categories.Count, layout);
            Func<double, double> toPixelY = value => plotTop + plotHeight - plotHeight * GetUnclampedPlotRatio(value, range.Min, range.Max);

            List<OfficePoint> points;
            string? equation = null;
            double? rSquared = null;
            if (trendline.Type == OfficeChartTrendlineType.MovingAverage) {
                int period = trendline.Period ?? 2;
                points = new List<OfficePoint>();
                for (int i = period - 1; i < ys.Count; i++) {
                    double average = 0D;
                    for (int j = i - period + 1; j <= i; j++) average += ys[j];
                    points.Add(new OfficePoint(toPixelX(xs[i]), toPixelY(average / period)));
                }
            } else {
                if (!TryFitTrendline(trendline, xs, ys, out Func<double, double>? model, out equation, out rSquared)) continue;
                double first = xs.Min();
                double last = xs.Max();
                points = new List<OfficePoint>(TrendlineSamples + 1);
                for (int i = 0; i <= TrendlineSamples; i++) {
                    double x = first + (last - first) * i / TrendlineSamples;
                    double y = model!(x);
                    if (IsFiniteChartValue(y)) points.Add(new OfficePoint(toPixelX(x), toPixelY(y)));
                }
            }

            if (points.Count < 2) continue;
            OfficeColor color = trendline.Color ?? GetSeriesColor(style, snapshot.Data.Series, seriesIndex);
            AddTrendlinePath(geometry, points, color, trendline.Width?.ToPoints() ?? 1.5D);

            var parts = new List<string>(2);
            if (trendline.DisplayEquation && equation != null) parts.Add(equation);
            if (trendline.DisplayRSquared && rSquared.HasValue) parts.Add("R² = " + rSquared.Value.ToString("0.####", CultureInfo.InvariantCulture));
            if (parts.Count > 0) {
                // Label the right-hand end, which is the last point unless the axis is reversed.
                OfficePoint end = points[points.Count - 1].X >= points[0].X ? points[points.Count - 1] : points[0];
                labels.Add((string.Join("   ", parts), end.X, end.Y));
            }
        }

        AddClippedPlotGeometry(drawing, geometry, new ChartPlotBounds(plotLeft, plotTop, plotWidth, plotHeight));
        foreach ((string text, double x, double y) in labels) {
            // Equations use plain chart text rather than the series' data-label styling and stay
            // inside the plot, right-aligned near the end of the line and clear of data labels.
            double fontSize = layout.DataLabelFontSize;
            var font = new OfficeFontInfo(string.IsNullOrWhiteSpace(layout.DataLabelFontFamily) ? style.FontFamily : layout.DataLabelFontFamily!, fontSize);
            OfficeTextMeasurer measurer = OfficeTextMeasurer.Create(font);
            double width = Math.Min(plotWidth, measurer.MeasureWidth(text, measurer.CreateStyle(font)) *
                OfficeTextMeasurer.PointsPerInch / OfficeTextMeasurer.DefaultDpi + 4D);
            double height = Math.Min(plotHeight, fontSize + 4D);
            double labelX = Math.Max(plotLeft, Math.Min(x - width, plotLeft + plotWidth - width));
            double labelY = Math.Max(plotTop, Math.Min(y - height - 4D, plotTop + plotHeight - height));
            labelY = AvoidDataLabelCollision(drawing, labelX, labelY, width, height, plotTop, plotTop + plotHeight);
            AddChartText(drawing, text, labelX, labelY, width, height, fontSize, style.TextColor, OfficeTextAlignment.Right, style,
                layout.DataLabelFontFamily);
        }
    }

    private static double CategoryNumberToPixel(OfficeChartSnapshot snapshot, double plotLeft, double plotWidth,
        double categoryNumber, int categoryCount, OfficeChartLayout layout) {
        if (categoryCount <= 1) return GetCartesianCategoryPointX(snapshot, plotLeft, plotWidth, 0, Math.Max(1, categoryCount), layout);
        double first = GetCartesianCategoryPointX(snapshot, plotLeft, plotWidth, 0, categoryCount, layout);
        double last = GetCartesianCategoryPointX(snapshot, plotLeft, plotWidth, categoryCount - 1, categoryCount, layout);
        return first + (last - first) * (categoryNumber - 1D) / (categoryCount - 1D);
    }

    private static void AddTrendlinePath(OfficeDrawing drawing, IReadOnlyList<OfficePoint> points, OfficeColor color, double width) {
        // Keep the path inside the drawing; the plot clip hides whatever lies beyond the axes.
        var clamped = points.Select(point => new OfficePoint(
            Math.Max(0D, Math.Min(drawing.Width, point.X)),
            Math.Max(0D, Math.Min(drawing.Height, point.Y)))).ToList();
        double left = clamped.Min(point => point.X);
        double top = clamped.Min(point => point.Y);
        var commands = new List<OfficePathCommand>(clamped.Count) { OfficePathCommand.MoveTo(clamped[0].X - left, clamped[0].Y - top) };
        for (int i = 1; i < clamped.Count; i++) commands.Add(OfficePathCommand.LineTo(clamped[i].X - left, clamped[i].Y - top));
        AddShape(drawing, OfficeShape.Path(commands), left, top, null, color, width, OfficeStrokeDashStyle.Solid);
    }

    private static bool TryFitTrendline(OfficeChartTrendline trendline, IReadOnlyList<double> xs, IReadOnlyList<double> ys,
        out Func<double, double>? model, out string? equation, out double? rSquared) {
        model = null;
        equation = null;
        rSquared = null;
        int n = xs.Count;
        switch (trendline.Type) {
            case OfficeChartTrendlineType.Linear: {
                    if (!TryFitPolynomial(xs, ys, 1, trendline.Intercept, out double[] c)) return false;
                    model = x => c[0] + c[1] * x;
                    equation = "y = " + FormatTerm(c[1], "x", leading: true) + FormatConstant(c[0]);
                    break;
                }
            case OfficeChartTrendlineType.Polynomial: {
                    int order = trendline.Order ?? 2;
                    if (n <= order || !TryFitPolynomial(xs, ys, order, trendline.Intercept, out double[] c)) return false;
                    model = x => { double sum = 0D, power = 1D; for (int k = 0; k < c.Length; k++) { sum += c[k] * power; power *= x; } return sum; };
                    var text = new System.Text.StringBuilder("y = ");
                    bool leading = true;
                    for (int k = order; k >= 1; k--) {
                        if (c[k] == 0D && k != order) continue;
                        text.Append(FormatTerm(c[k], k == 1 ? "x" : "x" + ToSuperscript(k), leading));
                        leading = false;
                    }
                    text.Append(FormatConstant(c[0]));
                    equation = text.ToString();
                    break;
                }
            case OfficeChartTrendlineType.Exponential: {
                    if (ys.Any(y => y <= 0D)) return false;
                    double? fixedLog = trendline.Intercept is double i && i > 0D ? Math.Log(i) : null;
                    if (!TryFitPolynomial(xs, ys.Select(value => Math.Log(value)).ToList(), 1, fixedLog, out double[] c)) return false;
                    double a = Math.Exp(c[0]), b = c[1];
                    model = x => a * Math.Exp(b * x);
                    equation = "y = " + FormatNumber(a) + "e" + FormatExponent(b, "x");
                    break;
                }
            case OfficeChartTrendlineType.Logarithmic: {
                    if (xs.Any(x => x <= 0D)) return false;
                    if (!TryFitPolynomial(xs.Select(value => Math.Log(value)).ToList(), ys, 1, null, out double[] c)) return false;
                    model = x => x > 0D ? c[0] + c[1] * Math.Log(x) : double.NaN;
                    equation = "y = " + FormatTerm(c[1], "ln(x)", leading: true) + FormatConstant(c[0]);
                    break;
                }
            case OfficeChartTrendlineType.Power: {
                    if (xs.Any(x => x <= 0D) || ys.Any(y => y <= 0D)) return false;
                    if (!TryFitPolynomial(xs.Select(value => Math.Log(value)).ToList(), ys.Select(value => Math.Log(value)).ToList(), 1, null, out double[] c)) return false;
                    double a = Math.Exp(c[0]), b = c[1];
                    model = x => x > 0D ? a * Math.Pow(x, b) : double.NaN;
                    equation = "y = " + FormatNumber(a) + "x" + FormatExponent(b, string.Empty);
                    break;
                }
            default:
                return false;
        }

        double mean = ys.Average();
        double total = 0D, residual = 0D;
        for (int i = 0; i < n; i++) {
            double predicted = model!(xs[i]);
            total += (ys[i] - mean) * (ys[i] - mean);
            residual += (ys[i] - predicted) * (ys[i] - predicted);
        }

        rSquared = total > 0D ? Math.Max(0D, 1D - residual / total) : 1D;
        return true;
    }

    // Least-squares polynomial fit; coefficients are returned lowest power first.
    private static bool TryFitPolynomial(IReadOnlyList<double> xs, IReadOnlyList<double> ys, int order, double? intercept, out double[] coefficients) {
        int start = intercept.HasValue ? 1 : 0;
        int size = order + 1 - start;
        coefficients = new double[order + 1];
        if (size <= 0) return false;
        var matrix = new double[size, size + 1];
        for (int i = 0; i < xs.Count; i++) {
            double target = ys[i] - (intercept ?? 0D);
            for (int row = 0; row < size; row++) {
                double rowPower = Math.Pow(xs[i], row + start);
                for (int column = 0; column < size; column++) matrix[row, column] += rowPower * Math.Pow(xs[i], column + start);
                matrix[row, size] += rowPower * target;
            }
        }

        for (int pivot = 0; pivot < size; pivot++) {
            int best = pivot;
            for (int row = pivot + 1; row < size; row++)
                if (Math.Abs(matrix[row, pivot]) > Math.Abs(matrix[best, pivot])) best = row;
            if (Math.Abs(matrix[best, pivot]) < 1E-12) return false;
            if (best != pivot)
                for (int column = 0; column <= size; column++) (matrix[pivot, column], matrix[best, column]) = (matrix[best, column], matrix[pivot, column]);
            for (int row = 0; row < size; row++) {
                if (row == pivot) continue;
                double factor = matrix[row, pivot] / matrix[pivot, pivot];
                for (int column = pivot; column <= size; column++) matrix[row, column] -= factor * matrix[pivot, column];
            }
        }

        for (int row = 0; row < size; row++) coefficients[row + start] = matrix[row, size] / matrix[row, row];
        if (intercept.HasValue) coefficients[0] = intercept.Value;
        return coefficients.All(IsFiniteChartValue);
    }

    private static string FormatNumber(double value) {
        if (Math.Abs(value) < 1E-12) return "0";
        string text = value.ToString("G4", CultureInfo.InvariantCulture);
        return text.Replace("E+0", "E").Replace("E-0", "E-");
    }

    private static string FormatTerm(double coefficient, string variable, bool leading) {
        string magnitude = FormatNumber(Math.Abs(coefficient)) + variable;
        if (leading) return (coefficient < 0D ? "-" : string.Empty) + magnitude;
        return (coefficient < 0D ? " - " : " + ") + magnitude;
    }

    private static string FormatConstant(double value) =>
        Math.Abs(value) < 1E-12 ? string.Empty : (value < 0D ? " - " : " + ") + FormatNumber(Math.Abs(value));

    // Exponents stay in characters every PDF base font can encode: squared and cubed use
    // the Latin-1 superscripts, other powers use a caret.
    private static string FormatExponent(double value, string variable) {
        string text = FormatNumber(value) + variable;
        return text.Length == 1 ? ToSuperscript(text) : "^(" + text + ")";
    }

    private static string ToSuperscript(int value) => ToSuperscript(value.ToString(CultureInfo.InvariantCulture));

    private static string ToSuperscript(string text) => text switch {
        "2" => "\u00B2",
        "3" => "\u00B3",
        _ => "^" + text
    };
}
