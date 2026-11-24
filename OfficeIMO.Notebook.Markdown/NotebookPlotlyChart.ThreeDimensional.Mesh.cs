using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaxMeshFaces = 4096;

    private static SceneTrace? ReadSceneMesh(JsonElement trace, JsonElement template, OfficeColor defaultColor, int remaining) {
        if (!ContainsOnly(trace, "type", "scene", "x", "y", "z", "i", "j", "k", "name", "color", "facecolor",
                "vertexcolor", "intensity", "intensitymode", "cmin", "cmax", "colorscale", "showscale", "reversescale",
                "colorbar", "opacity", "flatshading", "showlegend", "legendgroup", "hoverinfo", "hovertext", "hovertemplate",
                "text", "customdata", "meta", "visible") ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showlegend"), false, out bool showLegend) ||
            !TryReadBoolean(Property(trace, "flatshading"), false, out _) ||
            !TryReadBoolean(Property(trace, "showscale"), true, out bool showScale) ||
            !TryReadBoolean(Property(trace, "reversescale"), false, out bool reverse) ||
            !TryReadUnitInterval(Property(trace, "opacity"), 1, out double opacity) ||
            !TryReadColor(Property(trace, "color"), out OfficeColor? color) ||
            !TryReadOptionalText(Property(trace, "name"), out string? name)) return null;

        double[]? x = ReadNumbers(Property(trace, "x"), remaining);
        if (x == null || x.Length < 3) return null;
        double[]? y = ReadNumbers(Property(trace, "y"), x.Length);
        double[]? z = ReadNumbers(Property(trace, "z"), x.Length);
        if (y?.Length != x.Length || z?.Length != x.Length) return null;

        int[]? i = ReadMeshIndices(Property(trace, "i"), MaxMeshFaces, x.Length);
        if (i == null) return null;
        int[]? j = ReadMeshIndices(Property(trace, "j"), i.Length, x.Length);
        int[]? k = ReadMeshIndices(Property(trace, "k"), i.Length, x.Length);
        if (j?.Length != i.Length || k?.Length != i.Length) return null;

        JsonElement intensityMode = Property(trace, "intensitymode");
        bool perFace = ValueOrDefault(intensityMode, "cell", allowMissing: false);
        if (!perFace && !ValueOrDefault(intensityMode, "vertex")) return null;
        if (!TryReadMeshTriangleColors(trace, template, i.Length, x.Length, perFace, reverse, showScale,
                out OfficeColor[]? faceColors, out OfficeColor[]? vertexColors, out double[]? intensities, out SceneScale? scale))
            return null;
        string? scaleTitle = null;
        JsonElement colorbar = Property(trace, "colorbar");
        if (scale is { Show: true } &&
            ((colorbar.ValueKind is not (JsonValueKind.Undefined or JsonValueKind.Object)) ||
             !TryReadLayoutTitle(Property(colorbar, "title"), out scaleTitle))) return null;

        var points = new SceneVector[x.Length];
        for (int n = 0; n < points.Length; n++) points[n] = new SceneVector(x[n], y[n], z[n]);
        var faces = new List<SceneFace>(i.Length);
        for (int n = 0; n < i.Length; n++) {
            int a = i[n];
            int b = j[n];
            int c = k[n];
            if (a == b || b == c || a == c) continue;
            if (SceneVector.Cross(points[b] - points[a], points[c] - points[a]).Length == 0) continue;
            double intensity = intensities == null ? 0 : perFace ? intensities[n] : (intensities[a] + intensities[b] + intensities[c]) / 3;
            OfficeColor? faceColor = faceColors?[n];
            if (vertexColors != null) faceColor = AverageMeshColor(vertexColors[a], vertexColors[b], vertexColors[c]);
            faces.Add(new SceneFace(a, b, c, intensity, faceColor));
        }
        return faces.Count == 0
            ? null
            : new SceneTrace(points, false, false, color ?? defaultColor, default, showLegend, name, faces, scale,
                IsMesh: true, Opacity: opacity, ScaleTitle: scaleTitle);
    }

    private static int[]? ReadMeshIndices(JsonElement source, int maxCount, int pointCount) {
        double[]? values = ReadNumbers(source, maxCount);
        if (values == null) return null;
        var result = new int[values.Length];
        for (int i = 0; i < values.Length; i++) {
            double value = values[i];
            if (value < 0 || value >= pointCount || value != (int)value) return null;
            result[i] = (int)value;
        }
        return result;
    }

    private static bool TryReadMeshTriangleColors(JsonElement trace, JsonElement template, int faceCount, int pointCount,
        bool perFace, bool reverse, bool showScale, out OfficeColor[]? faceColors, out OfficeColor[]? vertexColors,
        out double[]? intensities, out SceneScale? scale) {
        scale = null;
        intensities = null;
        faceColors = null;
        vertexColors = null;
        // Match Plotly's selection order when the spec supplies more than one source.
        JsonElement intensitySource = Property(trace, "intensity");
        if (intensitySource.ValueKind != JsonValueKind.Undefined) {
            int count = perFace ? faceCount : pointCount;
            intensities = ReadNumbers(intensitySource, count);
            return intensities?.Length == count
                && TryReadSceneScale(trace, template, intensities, reverse, showScale, out scale);
        }
        JsonElement vertexColorSource = Property(trace, "vertexcolor");
        if (vertexColorSource.ValueKind != JsonValueKind.Undefined)
            return (vertexColors = ReadMeshColors(vertexColorSource, pointCount)) != null;
        JsonElement faceColorSource = Property(trace, "facecolor");
        return faceColorSource.ValueKind == JsonValueKind.Undefined
            || (faceColors = ReadMeshColors(faceColorSource, faceCount)) != null;
    }

    private static OfficeColor[]? ReadMeshColors(JsonElement source, int count) {
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() != count) return null;
        var result = new OfficeColor[count];
        int index = 0;
        foreach (JsonElement value in source.EnumerateArray()) {
            if (!TryReadColor(value, out OfficeColor? color) || !color.HasValue) return null;
            result[index++] = color.Value;
        }
        return result;
    }

    private static OfficeColor AverageMeshColor(OfficeColor a, OfficeColor b, OfficeColor c) =>
        OfficeColor.FromRgba((byte)((a.R + b.R + c.R) / 3), (byte)((a.G + b.G + c.G) / 3),
            (byte)((a.B + b.B + c.B) / 3), (byte)((a.A + b.A + c.A) / 3));

    private static OfficeColor ShadeMeshFace(SceneTrace trace, SceneFace face, SceneView view) {
        OfficeColor color = face.Color ?? trace.Scale?.At(face.ColorValue) ?? trace.SeriesColor;
        SceneVector a = view.Domain.Normalize(trace.Points[face.A]);
        SceneVector b = view.Domain.Normalize(trace.Points[face.B]);
        SceneVector c = view.Domain.Normalize(trace.Points[face.C]);
        SceneVector normal = SceneVector.Cross(b - a, c - a);
        SceneVector light = new SceneVector(0.30, -0.40, 0.8660254);
        // Use the same two-sided diffuse light regardless of the producer's triangle winding.
        double diffuse = normal.Length > 0 ? Math.Min(1, Math.Abs(SceneVector.Dot(normal, light)) / normal.Length) : 0;
        double strength = 0.56 + 0.44 * diffuse;
        return OfficeColor.FromRgba((byte)Math.Round(color.R * strength), (byte)Math.Round(color.G * strength),
            (byte)Math.Round(color.B * strength), (byte)Math.Round(color.A * trace.Opacity));
    }

    private static bool TryReadUnitInterval(JsonElement source, double fallback, out double result) {
        result = fallback;
        return source.ValueKind == JsonValueKind.Undefined || TryReadNumber(source, out result) && result >= 0 && result <= 1;
    }
}
