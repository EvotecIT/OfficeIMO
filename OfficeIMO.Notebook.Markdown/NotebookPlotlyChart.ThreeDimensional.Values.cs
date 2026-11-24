using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static SceneTrace? ReadScatter3D(JsonElement trace, OfficeColor seriesColor, int remaining) {
        if (!ContainsOnly(trace, "type", "scene", "x", "y", "z", "name", "mode", "marker", "line", "showlegend",
                "legendgroup", "hoverinfo", "hovertext", "hovertemplate", "customdata", "meta", "visible") ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showlegend"), true, out bool legend) ||
            !TryReadOptionalText(Property(trace, "name"), out string? name)) return null;

        double[]? x = ReadNumbers(Property(trace, "x"), remaining);
        if (x == null) return null;
        double[]? y = ReadNumbers(Property(trace, "y"), x.Length);
        double[]? z = ReadNumbers(Property(trace, "z"), x.Length);
        if (y == null || z == null || x.Length != y.Length || x.Length != z.Length ||
            !TryReadScatterMode(Property(trace, "mode"), x.Length, out bool markers, out bool lines) ||
            !TryReadTraceStyle(trace, bar: false, markers, lines, out PlotlyTraceStyle style) ||
            style.MarkerShape == OfficeChartMarkerShape.Triangle || (!markers && x.Length < 2)) return null;

        var points = new SceneVector[x.Length];
        for (int i = 0; i < points.Length; i++) points[i] = new SceneVector(x[i], y[i], z[i]);
        return new SceneTrace(points, markers, lines, seriesColor, style, legend, name, null, null);
    }

    private static SceneTrace? ReadSurface(JsonElement trace, JsonElement template, int remaining) {
        if (!ContainsOnly(trace, "type", "scene", "x", "y", "z", "surfacecolor", "colorscale", "cmin", "cmax",
                "showscale", "reversescale", "name", "showlegend", "hoverinfo", "hovertemplate", "hovertext", "meta", "visible") ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showscale"), true, out bool showScale) ||
            !TryReadBoolean(Property(trace, "reversescale"), false, out bool reverse) ||
            !TryReadBoolean(Property(trace, "showlegend"), false, out bool showLegend) || showLegend ||
            !TryReadGrid(Property(trace, "z"), remaining, out int rows, out int columns, out double[]? heights)) return null;

        double[]? x = ReadGridAxis(Property(trace, "x"), columns);
        double[]? y = ReadGridAxis(Property(trace, "y"), rows);
        if (x == null || y == null) return null;
        double[]? colors = heights;
        JsonElement colorData = Property(trace, "surfacecolor");
        if (colorData.ValueKind != JsonValueKind.Undefined &&
            (!TryReadGrid(colorData, remaining, out int colorRows, out int colorColumns, out colors) ||
                colorRows != rows || colorColumns != columns)) return null;
        if (!TryReadSceneScale(trace, template, colors!, reverse, showScale, out SceneScale? scale)) return null;

        var points = new SceneVector[rows * columns];
        var faces = new List<SceneFace>(2 * (rows - 1) * (columns - 1));
        for (int row = 0; row < rows; row++) {
            for (int column = 0; column < columns; column++) {
                int at = row * columns + column;
                points[at] = new SceneVector(x[column], y[row], heights![at]);
                if (row + 1 == rows || column + 1 == columns) continue;
                int nextRow = at + columns;
                faces.Add(CreateSurfaceFace(at, at + 1, nextRow + 1, colors!));
                faces.Add(CreateSurfaceFace(at, nextRow + 1, nextRow, colors!));
            }
        }
        return new SceneTrace(points, false, false, default, default, false, null, faces, scale);
    }

    private static SceneFace CreateSurfaceFace(int a, int b, int c, double[] colors) =>
        new SceneFace(a, b, c, (colors[a] + colors[b] + colors[c]) / 3);

    private static bool TryReadGrid(JsonElement source, int remaining, out int rows, out int columns, out double[]? values) {
        rows = 0;
        columns = 0;
        values = null;
        if (source.ValueKind == JsonValueKind.Object) {
            if (!TryReadString(Property(source, "shape"), 12, out string? shape)) return false;
            string[] dimensions = shape!.Split(',');
            if (dimensions.Length != 2 ||
                !int.TryParse(dimensions[0].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out rows) ||
                !int.TryParse(dimensions[1].Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out columns) ||
                !IsSupportedGridSize(rows, columns, remaining)) return false;
            values = ReadTypedNumbers(source, rows * columns, hasMatrixShape: true);
            return values?.Length == rows * columns;
        }
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() < 2) return false;
        rows = source.GetArrayLength();
        if (source[0].ValueKind != JsonValueKind.Array) return false;
        columns = source[0].GetArrayLength();
        if (!IsSupportedGridSize(rows, columns, remaining)) return false;
        values = new double[rows * columns];
        int index = 0;
        foreach (JsonElement row in source.EnumerateArray()) {
            if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() != columns) return false;
            foreach (JsonElement element in row.EnumerateArray()) {
                if (!TryReadNumber(element, out values[index++])) return false;
            }
        }
        return true;
    }

    private static bool IsSupportedGridSize(int rows, int columns, int remaining) =>
        rows is >= 2 and <= MaxSurfaceSide && columns is >= 2 and <= MaxSurfaceSide && rows * columns <= remaining;

    private static double[]? ReadGridAxis(JsonElement axis, int count) {
        if (axis.ValueKind == JsonValueKind.Undefined) return DefaultScatterX(count);
        double[]? values = ReadNumbers(axis, count);
        if (values?.Length != count || values[0] == values[1]) return null;
        bool increasing = values[1] > values[0];
        for (int i = 1; i < values.Length; i++) {
            if (increasing ? values[i] <= values[i - 1] : values[i] >= values[i - 1]) return null;
        }
        return values;
    }

    private static bool TryReadSceneScale(JsonElement trace, JsonElement template, double[] colors, bool reverse, bool showScale,
        out SceneScale? scale) {
        scale = null;
        JsonElement lower = Property(trace, "cmin");
        JsonElement upper = Property(trace, "cmax");
        double min = colors.Min();
        double max = colors.Max();
        if ((lower.ValueKind != JsonValueKind.Undefined || upper.ValueKind != JsonValueKind.Undefined) &&
            (!TryReadNumber(lower, out min) || !TryReadNumber(upper, out max) || min >= max)) return false;
        JsonElement source = WithTemplateDefault(trace, template, "colorscale");
        SceneColorStop[]? stops = ReadSceneColorStops(source);
        if (stops == null) return false;
        scale = new SceneScale(stops, min, max, reverse, showScale);
        return true;
    }

    private static SceneColorStop[]? ReadSceneColorStops(JsonElement source) {
        if (source.ValueKind == JsonValueKind.Undefined) {
            return new[] {
                new SceneColorStop(0, OfficeColor.FromRgb(68, 1, 84)),
                new SceneColorStop(0.5, OfficeColor.FromRgb(33, 145, 140)),
                new SceneColorStop(1, OfficeColor.FromRgb(253, 231, 37))
            };
        }
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is < 2 or > 32) return null;
        var stops = new SceneColorStop[source.GetArrayLength()];
        int index = 0;
        double previous = -1;
        foreach (JsonElement entry in source.EnumerateArray()) {
            if (entry.ValueKind != JsonValueKind.Array || entry.GetArrayLength() != 2 ||
                !TryReadNumber(entry[0], out double at) || at <= previous || at < 0 || at > 1 ||
                !TryReadColor(entry[1], out OfficeColor? color) || !color.HasValue || color.Value.A != 255) return null;
            stops[index++] = new SceneColorStop(at, color.Value);
            previous = at;
        }
        return stops[0].At == 0 && stops[stops.Length - 1].At == 1 ? stops : null;
    }

    private sealed record SceneTrace(SceneVector[] Points, bool Markers, bool Lines, OfficeColor SeriesColor,
        PlotlyTraceStyle Style, bool Legend, string? Name, List<SceneFace>? Faces, SceneScale? Scale);

    private readonly record struct SceneFace(int A, int B, int C, double ColorValue);

    private readonly record struct SceneColorStop(double At, OfficeColor Color);

    private sealed record SceneScale(SceneColorStop[] Stops, double Min, double Max, bool Reverse, bool Show) {
        internal OfficeColor At(double value) {
            double t = Max > Min ? Math.Max(0, Math.Min(1, (value - Min) / (Max - Min))) : 0.5;
            return FromFraction(t);
        }

        internal OfficeColor FromFraction(double t) {
            if (Reverse) t = 1 - t;
            for (int index = 1; index < Stops.Length; index++) {
                SceneColorStop before = Stops[index - 1];
                SceneColorStop after = Stops[index];
                if (t > after.At) continue;
                double weight = (t - before.At) / (after.At - before.At);
                return InterpolateSceneColor(before.Color, after.Color, weight);
            }
            return Stops[Stops.Length - 1].Color;
        }
    }
}
