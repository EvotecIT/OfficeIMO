using System;
using System.Linq;
using System.Globalization;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Uses percentage-only pie labels or count-only bar labels, and removes the chart background and frame.</summary>
    /// <param name="percentageLabels">True for percentage-only pie labels; false for numeric count labels.</param>
    /// <param name="titleColor">RGB hex text color for the chart title.</param>
    /// <returns>The chart.</returns>
    public WordChart SetPaperStyle(bool percentageLabels = false, string titleColor = "202428") {
        var root = _chartPart?.ChartSpace;
        if (root == null) return this;
        C.ShapeProperties? frame = root.GetFirstChild<C.ShapeProperties>();
        if (frame == null) {
            frame = new C.ShapeProperties();
            root.AddChild(frame, true);
        }
        frame.RemoveAllChildren<A.SolidFill>();
        frame.RemoveAllChildren<A.GradientFill>();
        frame.RemoveAllChildren<A.PatternFill>();
        frame.RemoveAllChildren<A.BlipFill>();
        frame.RemoveAllChildren<A.GroupFill>();
        frame.RemoveAllChildren<A.NoFill>();
        frame.AddChild(new A.NoFill(), true);
        frame.RemoveAllChildren<A.Outline>();
        frame.AddChild(new A.Outline(new A.NoFill()), true);
        foreach (A.DefaultRunProperties properties in root.Descendants<C.Title>().SelectMany(title => title.Descendants<A.DefaultRunProperties>())) {
            properties.RemoveAllChildren<A.SolidFill>();
            properties.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = titleColor.TrimStart('#') }), true);
        }
        foreach (C.DataLabels labels in root.Descendants<C.DataLabels>()) {
            if (!percentageLabels) {
                labels.RemoveAllChildren<C.DataLabelPosition>();
                labels.AddChild(new C.DataLabelPosition { Val = C.DataLabelPositionValues.OutsideEnd }, true);
            }
            labels.RemoveAllChildren<C.ShowValue>();
            labels.RemoveAllChildren<C.ShowPercent>();
            labels.RemoveAllChildren<C.NumberingFormat>();
            labels.AddChild(new C.ShowValue { Val = !percentageLabels }, true);
            labels.AddChild(new C.ShowPercent { Val = percentageLabels }, true);
            labels.AddChild(new C.NumberingFormat { FormatCode = percentageLabels ? "0.#%" : "0", SourceLinked = false }, true);
        }
        root.Save();
        return this;
    }

    /// <summary>Sets count axes to start at zero, use whole-number labels and a positive integral tick interval.</summary>
    /// <param name="majorUnit">Optional distance between ticks in whole counts. When omitted, a readable integral interval is selected from the cached values.</param>
    /// <returns>The chart.</returns>
    public WordChart SetCountAxis(uint? majorUnit = null) {
        if (majorUnit == 0) throw new ArgumentOutOfRangeException(nameof(majorUnit));
        var root = _chartPart?.ChartSpace;
        if (root == null) return this;
        double interval = majorUnit.HasValue ? majorUnit.Value : SelectCountAxisInterval(root);
        foreach (C.ValueAxis axis in root.Descendants<C.ValueAxis>()) {
            C.Scaling? scaling = axis.GetFirstChild<C.Scaling>();
            if (scaling == null) {
                scaling = new C.Scaling();
                axis.AddChild(scaling, true);
            }
            scaling.MinAxisValue = new C.MinAxisValue { Val = 0 };
            double dataMaximum = ReadCountMaximum(root, includeAuthoredMaximum: false);
            double requiredMaximum = Math.Ceiling(dataMaximum / interval + 1D) * interval;
            scaling.MaxAxisValue = new C.MaxAxisValue { Val = Math.Max(scaling.MaxAxisValue?.Val?.Value ?? 0D, requiredMaximum) };
            axis.RemoveAllChildren<C.NumberingFormat>();
            axis.AddChild(new C.NumberingFormat { FormatCode = "0", SourceLinked = false }, true);
            // MajorUnit is a schema child without a generated CLR property in OpenXML 3.5.1.
            axis.RemoveAllChildren<C.MajorUnit>();
            axis.AddChild(new C.MajorUnit { Val = interval }, true);
        }
        root.Save();
        return this;
    }

    private static double SelectCountAxisInterval(C.ChartSpace root) {
        double maximum = ReadCountMaximum(root, includeAuthoredMaximum: true);
        // Allow the padded upper bound to retain its interval when formatting is repeated.
        // Data targets ten intervals; authored bounds may use up to twelve, below the reader's limit.
        double dataMaximum = ReadCountMaximum(root, includeAuthoredMaximum: false);
        double target = Math.Max(1D, Math.Max(dataMaximum / 10D, maximum / 12D));
        double scale = Math.Pow(10D, Math.Floor(Math.Log10(target)));
        double normalized = target / scale;
        return Math.Max(1D, (normalized <= 1D ? 1D : normalized <= 2D ? 2D : normalized <= 5D ? 5D : 10D) * scale);
    }

    private static double ReadCountMaximum(C.ChartSpace root, bool includeAuthoredMaximum) {
        double Read(C.NumericPoint point) => double.TryParse(point.NumericValue?.Text, NumberStyles.Float,
            CultureInfo.InvariantCulture, out double value) && !double.IsNaN(value) && !double.IsInfinity(value) ? Math.Max(0D, value) : 0D;
        double maximum = root.Descendants<C.Values>().SelectMany(values => values.Descendants<C.NumericPoint>())
            .Select(Read).DefaultIfEmpty(0D).Max();
        if (includeAuthoredMaximum) maximum = Math.Max(maximum, root.Descendants<C.MaxAxisValue>()
            .Select(bound => bound.Val?.Value ?? 0D).DefaultIfEmpty(0D).Max());
        // Stacked bars use the sum at each category, rather than an individual series maximum.
        foreach (C.BarChart bar in root.Descendants<C.BarChart>().Where(bar => bar.BarGrouping?.Val?.Value == C.BarGroupingValues.Stacked)) {
            double stacked = bar.Elements<C.BarChartSeries>().SelectMany(series => series.GetFirstChild<C.Values>()?.Descendants<C.NumericPoint>()
                    ?? Enumerable.Empty<C.NumericPoint>())
                .GroupBy(point => point.Index?.Value ?? 0U).Select(points => points.Sum(Read)).DefaultIfEmpty(0D).Max();
            maximum = Math.Max(maximum, stacked);
        }
        return maximum;
    }
}
