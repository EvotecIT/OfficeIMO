using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaxSceneTraces = 8;
    private const int MaxSurfaceSide = 32;

    /// <summary>Projects bounded saved Plotly scatter, surface, or explicit mesh geometry through the saved camera.</summary>
    internal static OfficeDrawing? TryRenderThreeDimensional(JsonElement spec, CancellationToken cancellationToken, out string? title,
        List<NotebookChartMark>? marks = null) {
        title = null;
        cancellationToken.ThrowIfCancellationRequested();
        if (!ContainsOnly(spec, "data", "layout", "config")) return null;
        JsonElement data = Property(spec, "data");
        if (data.ValueKind != JsonValueKind.Array || data.GetArrayLength() is < 1 or > MaxSceneTraces) return null;
        if (!TryReadString(Property(data[0], "type"), 9, out string? kind) ||
            kind is not ("surface" or "mesh3d" or "scatter3d") ||
            (kind != "scatter3d" && data.GetArrayLength() != 1) ||
            !TryReadSceneLayout(Property(spec, "layout"), kind, data.GetArrayLength(), out SceneLayout? layout))
            return null;

        var traces = new List<SceneTrace>(data.GetArrayLength());
        int remaining = MaxPoints;
        foreach (JsonElement trace in data.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ValueOrDefault(Property(trace, "type"), kind, allowMissing: false) ||
                !ValueOrDefault(Property(trace, "scene"), "scene")) return null;
            OfficeColor seriesColor = layout!.Style.GetSeriesColor(traces.Count);
            SceneTrace? parsed = kind switch {
                "surface" => ReadSurface(trace, layout.TraceTemplate, remaining),
                "mesh3d" => ReadSceneMesh(trace, layout.TraceTemplate, seriesColor, remaining),
                _ => ReadScatter3D(trace, seriesColor, remaining)
            };
            if (parsed == null) return null;
            traces.Add(parsed);
            remaining -= parsed.Points.Length;
        }

        OfficeDrawing? drawing = RenderScene(layout!, traces, cancellationToken, marks);
        if (drawing != null) title = layout!.Title;
        return drawing;
    }

    private static bool TryReadSceneLayout(JsonElement source, string kind, int traceCount, out SceneLayout? layout) {
        layout = null;
        JsonElement template = Property(source, "template");
        JsonElement scene = Property(source, "scene");
        if (!ContainsOnly(source, "template", "title", "width", "height", "scene", "showlegend", "legend",
                "paper_bgcolor", "colorway", "font", "margin", "hovermode", "hoverlabel", "autosize") ||
            !TryReadLayoutTitle(Property(source, "title"), out string? title) ||
            !TryReadLength(Property(source, "width"), 320, 1600, out OfficeLength? width) ||
            !TryReadLength(Property(source, "height"), 240, 1200, out OfficeLength? height) ||
            !TryReadBoolean(Property(source, "showlegend"), traceCount > 1 && kind == "scatter3d", out bool showLegend) ||
            !TryReadLegend(Property(source, "legend"), out OfficeChartLegendPosition legend) ||
            !TryReadTemplate(template, "scatter", out JsonElement templateLayout) ||
            !TryReadSceneTemplate(template, kind, out JsonElement traceTemplate) ||
            !TryReadCamera(Property(scene, "camera"), out SceneCamera camera) ||
            !TryReadSceneAspect(scene, out SceneAspect aspect, out SceneVector aspectRatio) ||
            !TryReadSceneAxis(Property(scene, "xaxis"), "x", out string? xTitle) ||
            !TryReadSceneAxis(Property(scene, "yaxis"), "y", out string? yTitle) ||
            !TryReadSceneAxis(Property(scene, "zaxis"), "z", out string? zTitle) ||
            !TryReadChartStyle(source, templateLayout, out OfficeChartStyle? style) ||
            !TryReadSceneFont(WithTemplateDefault(source, templateLayout, "font"), out OfficeColor? textColor)) return false;

        layout = new SceneLayout(title, (width ?? OfficeLength.Pixels(700)).ToPoints(), (height ?? OfficeLength.Pixels(450)).ToPoints(),
            style!, textColor ?? style!.TextColor, camera, aspect, aspectRatio,
            xTitle, yTitle, zTitle, showLegend, legend, traceTemplate);
        return true;
    }

    private static bool TryReadSceneFont(JsonElement source, out OfficeColor? color) {
        color = null;
        return ContainsOnly(source, "color", "family", "size") && TryReadColor(Property(source, "color"), out color);
    }

    private static bool TryReadSceneTemplate(JsonElement template, string kind, out JsonElement trace) {
        trace = default;
        JsonElement scene = Property(Property(template, "layout"), "scene");
        if (!ContainsOnly(scene, "xaxis", "yaxis", "zaxis") ||
            !IsSceneAxisStyle(Property(scene, "xaxis")) ||
            !IsSceneAxisStyle(Property(scene, "yaxis")) ||
            !IsSceneAxisStyle(Property(scene, "zaxis"))) return false;

        JsonElement traces = Property(Property(template, "data"), kind);
        if (traces.ValueKind == JsonValueKind.Undefined) return true;
        if (traces.ValueKind != JsonValueKind.Array || traces.GetArrayLength() != 1) return false;
        trace = traces[0];
        if (!ValueOrDefault(Property(trace, "type"), kind, allowMissing: false)) return false;
        if (kind is "surface" or "mesh3d") {
            return ContainsOnly(trace, "type", "colorbar", "colorscale") &&
                ContainsOnly(Property(trace, "colorbar"), "outlinewidth", "ticks");
        }

        return ContainsOnly(trace, "type", "line", "marker") &&
            IsSceneTraceStyleTemplate(Property(trace, "line")) &&
            IsSceneTraceStyleTemplate(Property(trace, "marker"));
    }

    private static bool IsSceneAxisStyle(JsonElement value) =>
        ContainsOnly(value, "backgroundcolor", "gridcolor", "linecolor", "showbackground", "ticks", "zerolinecolor", "gridwidth");

    private static bool IsSceneTraceStyleTemplate(JsonElement value) =>
        ContainsOnly(value, "colorbar") && ContainsOnly(Property(value, "colorbar"), "outlinewidth", "ticks");

    private static bool TryReadSceneAxis(JsonElement value, string defaultTitle, out string? title) {
        title = defaultTitle;
        if (!ContainsOnly(value, "title", "type", "backgroundcolor", "gridcolor", "linecolor", "showbackground",
                "ticks", "zerolinecolor", "gridwidth") ||
            !ValueOrDefault(Property(value, "type"), "linear")) return false;
        JsonElement label = Property(value, "title");
        return label.ValueKind == JsonValueKind.Undefined || TryReadLayoutTitle(label, out title);
    }

    private static bool TryReadSceneAspect(JsonElement scene, out SceneAspect aspect, out SceneVector ratio) {
        aspect = SceneAspect.Auto;
        ratio = new SceneVector(1, 1, 1);
        if (!ContainsOnly(scene, "camera", "aspectmode", "aspectratio", "xaxis", "yaxis", "zaxis", "dragmode", "hovermode")) return false;
        JsonElement value = Property(scene, "aspectmode");
        JsonElement requestedRatio = Property(scene, "aspectratio");
        if (value.ValueKind == JsonValueKind.Undefined && requestedRatio.ValueKind != JsonValueKind.Undefined) {
            aspect = SceneAspect.Manual;
        } else if (value.ValueKind != JsonValueKind.Undefined) {
            if (value.ValueKind != JsonValueKind.String) return false;
            if (value.ValueEquals("cube")) aspect = SceneAspect.Cube;
            else if (value.ValueEquals("data")) aspect = SceneAspect.Data;
            else if (value.ValueEquals("manual")) aspect = SceneAspect.Manual;
            else if (!value.ValueEquals("auto")) return false;
        }
        if (aspect != SceneAspect.Manual) return requestedRatio.ValueKind == JsonValueKind.Undefined;
        return TryReadSceneVector(requestedRatio, ratio, 16, out ratio) &&
            ratio.X >= 0.05 && ratio.Y >= 0.05 && ratio.Z >= 0.05;
    }

    private static bool TryReadCamera(JsonElement camera, out SceneCamera result) {
        result = default;
        JsonElement projection = Property(camera, "projection");
        if (!ContainsOnly(camera, "eye", "center", "up", "projection") || !ContainsOnly(projection, "type") ||
            !TryReadSceneVector(Property(camera, "eye"), new SceneVector(1.25, 1.25, 1.25), 12, out SceneVector eye) ||
            !TryReadSceneVector(Property(camera, "center"), default, 0.5, out SceneVector center) ||
            !TryReadSceneVector(Property(camera, "up"), new SceneVector(0, 0, 1), 12, out SceneVector up)) return false;
        JsonElement projectionType = Property(projection, "type");
        bool orthographic = ValueOrDefault(projectionType, "orthographic", allowMissing: false);
        if (!orthographic && !ValueOrDefault(projectionType, "perspective")) return false;
        return SceneCamera.TryCreate(eye, center, up, orthographic, out result);
    }

    private static bool TryReadSceneVector(JsonElement element, SceneVector defaultValue, double max, out SceneVector result) {
        result = defaultValue;
        if (!ContainsOnly(element, "x", "y", "z") ||
            !TryReadVectorComponent(element, "x", defaultValue.X, max, out double x) ||
            !TryReadVectorComponent(element, "y", defaultValue.Y, max, out double y) ||
            !TryReadVectorComponent(element, "z", defaultValue.Z, max, out double z)) return false;
        result = new SceneVector(x, y, z);
        return true;
    }

    private static bool TryReadVectorComponent(JsonElement element, string name, double defaultValue, double max, out double value) {
        value = defaultValue;
        JsonElement component = Property(element, name);
        return component.ValueKind == JsonValueKind.Undefined || TryReadNumber(component, out value) && Math.Abs(value) <= max;
    }

    private enum SceneAspect { Auto, Cube, Data, Manual }

    private sealed record SceneLayout(string? Title, double Width, double Height, OfficeChartStyle Style,
        OfficeColor TextColor, SceneCamera Camera, SceneAspect Aspect, SceneVector Ratio, string? XTitle, string? YTitle,
        string? ZTitle, bool ShowLegend, OfficeChartLegendPosition Legend, JsonElement TraceTemplate);
}
