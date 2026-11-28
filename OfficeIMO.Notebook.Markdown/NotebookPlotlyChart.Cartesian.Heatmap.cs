using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaxHeatSide = 128;
    private const int MaxHeatCells = 4096;

    private static Heatmap? ReadHeatmap(CartesianPanel panel, CartesianLayout appearance, int remainingCells,
        HashSet<string> colorAxes, CancellationToken cancellationToken) {
        JsonElement trace = panel.Traces[0].Source;
        if (!ContainsOnly(trace, "type", "x", "y", "z", "xaxis", "yaxis", "colorscale", "coloraxis", "colorbar",
                "zmin", "zmax", "zauto", "name", "showlegend", "showscale", "reversescale", "xgap", "ygap", "hoverinfo",
                "hovertemplate", "hovertext", "hoverongaps", "customdata", "meta", "visible", "zsmooth", "texttemplate",
                "x0", "dx", "y0", "dy") ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showlegend"), false, out bool showInLegend) || showInLegend ||
            !TryReadBoolean(Property(trace, "zsmooth"), false, out bool smooth) || smooth ||
            !TryReadLength(Property(trace, "xgap"), 0, 32, out OfficeLength? xGap) ||
            !TryReadLength(Property(trace, "ygap"), 0, 32, out OfficeLength? yGap) ||
            !TryReadHeatmapLabel(Property(trace, "texttemplate"), out string? labelFormat)) return null;

        double[]? values = ReadHeatMatrix(Property(trace, "z"), remainingCells, cancellationToken, out int rows, out int columns);
        if (values == null ||
            !TryReadHeatAxis(trace, "x", columns, panel.X, out HeatAxis? x) ||
            !TryReadHeatAxis(trace, "y", rows, panel.Y, out HeatAxis? y) ||
            !TryReadHeatScale(trace, appearance, colorAxes, values, out SceneScale? scale, out string? scaleTitle)) return null;
        return new Heatmap(values, x!, y!, scale!, scaleTitle, xGap?.ToPoints() ?? 0, yGap?.ToPoints() ?? 0, labelFormat);
    }

    private static double[]? ReadHeatMatrix(JsonElement source, int remaining, CancellationToken cancellationToken,
        out int rows, out int columns) {
        rows = columns = 0;
        if (source.ValueKind == JsonValueKind.Object) {
            if (!TryReadString(Property(source, "shape"), 12, out string? shape)) return null;
            string[] parts = shape!.Split(',');
            if (parts.Length != 2 ||
                !int.TryParse(parts[0].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out rows) ||
                !int.TryParse(parts[1].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out columns) ||
                !ValidHeatSize(rows, columns, remaining)) return null;
            double[]? typed = ReadTypedNumbers(source, rows * columns, hasMatrixShape: true, allowNaN: true);
            return typed?.Length == rows * columns ? typed : null;
        }
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is < 1 or > MaxHeatSide ||
            source[0].ValueKind != JsonValueKind.Array) return null;
        rows = source.GetArrayLength();
        columns = source[0].GetArrayLength();
        if (!ValidHeatSize(rows, columns, remaining)) return null;
        var matrix = new double[rows * columns];
        int index = 0;
        foreach (JsonElement row in source.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() != columns) return null;
            foreach (JsonElement sample in row.EnumerateArray()) {
                if (sample.ValueKind == JsonValueKind.Null) matrix[index++] = double.NaN;
                else if (!TryReadNumber(sample, out matrix[index++])) return null;
            }
        }
        return matrix;
    }

    private static bool ValidHeatSize(int rows, int columns, int remaining) =>
        rows is >= 1 and <= MaxHeatSide && columns is >= 1 and <= MaxHeatSide && rows * columns <= remaining;

    private static bool TryReadHeatAxis(JsonElement trace, string key, int count, CartesianAxis axis, out HeatAxis? result) {
        result = null;
        JsonElement source = Property(trace, key);
        string[]? labels = null;
        double[]? values;
        double singletonInterval = 1;
        bool coordinatesPresent = source.ValueKind != JsonValueKind.Undefined;
        if (!coordinatesPresent) {
            if (!TryReadHeatCoordinates(trace, key, count, out values, out singletonInterval)) return false;
        } else {
            if (HasProperty(trace, key + "0") || HasProperty(trace, "d" + key)) return false;
            if (!axis.Linear && TryReadCategories(source, count, axis.Categorical, out labels)) {
                if (labels!.Length != count || axis.Minimum.HasValue) return false;
                values = DefaultScatterX(count);
            } else {
                if (axis.Categorical) return false;
                values = ReadNumbers(source, count + 1);
            }
        }
        if (values == null || (values.Length != count && values.Length != count + 1) || !MonotonicHeatAxis(values)) return false;
        double[] edges;
        double[] centers;
        if (values.Length == count + 1) {
            edges = values;
            centers = new double[count];
            for (int index = 0; index < count; index++) centers[index] = (edges[index] + edges[index + 1]) / 2;
        } else {
            centers = values;
            edges = GetHeatEdges(centers, singletonInterval);
        }
        double minimum = axis.Minimum ?? Math.Min(edges[0], edges[count]);
        double maximum = axis.Maximum ?? Math.Max(edges[0], edges[count]);
        if (maximum <= minimum) return false;
        result = new HeatAxis(centers, edges, labels, minimum, maximum, axis.Reversed);
        return true;
    }

    private static bool TryReadHeatCoordinates(JsonElement trace, string axis, int count, out double[]? values, out double step) {
        values = null;
        JsonElement sourceStart = Property(trace, axis + "0");
        JsonElement sourceStep = Property(trace, "d" + axis);
        double start = 0;
        step = 1;
        if ((sourceStart.ValueKind != JsonValueKind.Undefined && !TryReadNumber(sourceStart, out start)) ||
            (sourceStep.ValueKind != JsonValueKind.Undefined && (!TryReadNumber(sourceStep, out step) || step == 0)) ||
            !IsSupportedNumber(start + (count - 1) * step)) return false;
        values = new double[count];
        for (int index = 0; index < count; index++) values[index] = start + index * step;
        return true;
    }

    private static bool MonotonicHeatAxis(double[] values) {
        if (values.Length == 1) return true;
        bool increasing = values[1] > values[0];
        for (int index = 1; index < values.Length; index++) {
            if (increasing ? values[index] <= values[index - 1] : values[index] >= values[index - 1]) return false;
        }
        return true;
    }

    private static double[] GetHeatEdges(double[] centers, double singletonInterval) {
        if (centers.Length == 1) return new[] { centers[0] - singletonInterval / 2, centers[0] + singletonInterval / 2 };
        var edges = new double[centers.Length + 1];
        for (int index = 1; index < centers.Length; index++) edges[index] = (centers[index - 1] + centers[index]) / 2;
        edges[0] = 2 * centers[0] - edges[1];
        edges[centers.Length] = 2 * centers[centers.Length - 1] - edges[centers.Length - 1];
        return edges;
    }

    private static bool TryReadHeatScale(JsonElement trace, CartesianLayout appearance, HashSet<string> colorAxes, double[] values,
        out SceneScale? result, out string? title) {
        result = null;
        title = null;
        JsonElement sharedSource = Property(trace, "coloraxis");
        bool shared = sharedSource.ValueKind != JsonValueKind.Undefined;
        JsonElement scaleSource = trace;
        JsonElement traceDefaults = Property(Property(appearance.Template, "data"), "heatmap");
        if (traceDefaults.ValueKind == JsonValueKind.Array && traceDefaults.GetArrayLength() > 1) return false;
        JsonElement defaults = traceDefaults.ValueKind == JsonValueKind.Array && traceDefaults.GetArrayLength() == 1
            ? traceDefaults[0] : default;
        if (shared) {
            if (!TryReadCartesianId(sharedSource, "coloraxis", MaxCartesianPanels, out string? colorAxis) ||
                !colorAxes.Add(colorAxis!)) return false;
            scaleSource = Property(appearance.Source, colorAxis!);
            defaults = Property(Property(appearance.Template, "layout"), colorAxis!);
            if (!ContainsOnly(scaleSource, "colorscale", "colorbar", "cmin", "cmax", "cauto", "showscale", "reversescale") ||
                !ContainsOnly(defaults, "colorscale", "colorbar")) return false;
        }

        string minProperty = shared ? "cmin" : "zmin";
        string maxProperty = shared ? "cmax" : "zmax";
        string autoProperty = shared ? "cauto" : "zauto";
        JsonElement lower = Property(scaleSource, minProperty);
        JsonElement upper = Property(scaleSource, maxProperty);
        double min = double.PositiveInfinity;
        double max = double.NegativeInfinity;
        foreach (double value in values) {
            if (double.IsNaN(value)) continue;
            min = Math.Min(min, value);
            max = Math.Max(max, value);
        }
        if (min > max) return false;
        bool explicitBounds = lower.ValueKind != JsonValueKind.Undefined || upper.ValueKind != JsonValueKind.Undefined;
        if (!TryReadBoolean(Property(scaleSource, autoProperty), !explicitBounds, out bool automatic) ||
            (!automatic && (!TryReadNumber(lower, out min) || !TryReadNumber(upper, out max) || min >= max)) ||
            !TryReadBoolean(Property(scaleSource, "showscale"), true, out bool show) ||
            !TryReadBoolean(Property(scaleSource, "reversescale"), false, out bool reverse) ||
            !TryReadHeatColorbar(WithTemplateDefault(scaleSource, defaults, "colorbar"), out title)) return false;
        SceneColorStop[]? stops = ReadSceneColorStops(WithTemplateDefault(scaleSource, defaults, "colorscale"));
        if (stops == null) return false;
        result = new SceneScale(stops, min, max, reverse, show);
        return true;
    }

    private static bool TryReadHeatColorbar(JsonElement source, out string? title) {
        title = null;
        return ContainsOnly(source, "title", "outlinewidth", "ticks") &&
            TryReadLayoutTitle(Property(source, "title"), out title);
    }

    private static bool TryReadHeatmapLabel(JsonElement source, out string? format) {
        format = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(source, 9, out string? text)) return false;
        if (text == "%{z}") {
            format = "G6";
            return true;
        }
        if (text!.Length != 8 || !text.StartsWith("%{z:.", StringComparison.Ordinal) ||
            text[5] < '0' || text[5] > '6' || text[6] != 'f' || text[7] != '}') return false;
        format = "F" + text[5];
        return true;
    }

    private sealed record Heatmap(double[] Values, HeatAxis X, HeatAxis Y, SceneScale Colors, string? ColorbarTitle,
        double HorizontalGap, double VerticalGap, string? LabelFormat);
    private sealed record HeatAxis(double[] Centers, double[] Edges, string[]? Labels, double Minimum, double Maximum, bool Reversed) {
        internal string Label(int index, bool precise = false) => Labels?[index] ??
            Centers[index].ToString(precise ? "G" : "G4", CultureInfo.InvariantCulture);
        internal double Fraction(double value) => Reversed ? 1 - (value - Minimum) / (Maximum - Minimum) :
            (value - Minimum) / (Maximum - Minimum);
    }
}
