using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static bool TryReadTraceStyle(JsonElement trace, bool bar, bool markers, bool connect, out PlotlyTraceStyle style) {
        style = default;
        JsonElement marker = Property(trace, "marker");
        JsonElement line = Property(trace, "line");
        if (!ContainsOnly(marker, "color", "size", "symbol", "pattern") ||
            !ContainsOnly(line, "color", "width", "dash", "shape") ||
            !TryReadColor(Property(marker, "color"), out OfficeColor? markerColor) ||
            !TryReadColor(Property(line, "color"), out OfficeColor? lineColor) ||
            !TryReadLength(Property(marker, "size"), 0.5, 64, out OfficeLength? markerSize) ||
            !TryReadLength(Property(line, "width"), 0.5, 24, out OfficeLength? lineWidth) ||
            !TryReadMarkerShape(Property(marker, "symbol"), out OfficeChartMarkerShape? markerShape) ||
            !TryReadLineDash(Property(line, "dash"), out OfficeStrokeDashStyle? lineDash) ||
            !ValueOrDefault(Property(line, "shape"), "linear")) return false;

        // Plotly Express includes a blank pattern on a plain bar. Nonempty hatch patterns need their own mapping.
        JsonElement pattern = Property(marker, "pattern");
        if (!ContainsOnly(pattern, "shape") || !ValueOrDefault(Property(pattern, "shape"), string.Empty) ||
            (!bar && HasProperty(marker, "pattern")) ||
            (!markers && (markerSize.HasValue || markerShape.HasValue)) ||
            (!connect && line.ValueKind != JsonValueKind.Undefined)) return false;

        style = new PlotlyTraceStyle(markerColor, lineColor, markerSize, lineWidth, markerShape, lineDash);
        return true;
    }

    private static bool TryReadMarkerShape(JsonElement source, out OfficeChartMarkerShape? markerShape) {
        markerShape = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(source, 16, out string? symbol)) return false;
        markerShape = symbol switch {
            "circle" => OfficeChartMarkerShape.Circle,
            "square" => OfficeChartMarkerShape.Square,
            "diamond" => OfficeChartMarkerShape.Diamond,
            "triangle-up" => OfficeChartMarkerShape.Triangle,
            "x" => OfficeChartMarkerShape.X,
            _ => null
        };
        return markerShape.HasValue;
    }

    private static bool TryReadLineDash(JsonElement source, out OfficeStrokeDashStyle? dash) {
        dash = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(source, 12, out string? value)) return false;
        dash = value switch {
            "solid" => OfficeStrokeDashStyle.Solid,
            "dot" => OfficeStrokeDashStyle.Dot,
            "dash" => OfficeStrokeDashStyle.Dash,
            "dashdot" => OfficeStrokeDashStyle.DashDot,
            _ => null
        };
        return dash.HasValue;
    }

    private static bool TryReadTemplate(JsonElement template, string traceType, out JsonElement layout) {
        layout = Property(template, "layout");
        // A producer template can set defaults that change this Cartesian chart. Limit the
        // applicable layout and trace defaults to presentation fields and inactive styles.
        if (!ContainsOnly(template, "layout", "data") ||
            !ContainsOnly(layout, "annotationdefaults", "autotypenumbers", "coloraxis", "colorscale", "colorway",
                "font", "geo", "hoverlabel", "hovermode", "mapbox", "paper_bgcolor", "plot_bgcolor", "polar",
                "scene", "shapedefaults", "ternary", "title", "xaxis", "yaxis") ||
            !TryReadTemplateAxis(Property(layout, "xaxis")) ||
            !TryReadTemplateAxis(Property(layout, "yaxis")) ||
            !ContainsOnly(Property(layout, "title"), "x", "y", "xanchor", "yanchor", "font", "pad")) return false;

        JsonElement data = Property(template, "data");
        if (data.ValueKind == JsonValueKind.Undefined) return true;
        if (data.ValueKind != JsonValueKind.Object) return false;
        JsonElement traces = Property(data, traceType);
        if (traces.ValueKind == JsonValueKind.Undefined) return true;
        if (traces.ValueKind != JsonValueKind.Array || traces.GetArrayLength() > MaxTraces) return false;
        foreach (JsonElement trace in traces.EnumerateArray()) {
            if (trace.ValueKind != JsonValueKind.Object ||
                !ValueOrDefault(Property(trace, "type"), traceType) ||
                !TryReadTemplateTrace(trace, traceType)) return false;
        }
        return true;
    }

    private static bool TryReadTemplateAxis(JsonElement axis) =>
        ContainsOnly(axis, "automargin", "gridcolor", "linecolor", "ticks", "title", "zerolinecolor", "zerolinewidth") &&
        ContainsOnly(Property(axis, "title"), "standoff", "font");

    private static bool TryReadTemplateTrace(JsonElement trace, string traceType) {
        JsonElement marker = Property(trace, "marker");
        if (traceType == "scattergl") return ContainsOnly(trace, "type", "marker")
            && ContainsOnly(marker, "colorbar")
            && ContainsOnly(Property(marker, "colorbar"), "outlinewidth", "ticks");
        if (traceType != "bar") return ContainsOnly(trace, "type", "fillpattern");
        return ContainsOnly(trace, "type", "error_x", "error_y", "marker") &&
            ContainsOnly(Property(trace, "error_x"), "color") &&
            ContainsOnly(Property(trace, "error_y"), "color") &&
            ContainsOnly(marker, "line", "pattern") &&
            ContainsOnly(Property(marker, "line"), "color", "width") &&
            ContainsOnly(Property(marker, "pattern"), "fillmode", "size", "solidity");
    }

    private static bool TryReadChartStyle(JsonElement layout, JsonElement templateLayout, out OfficeChartStyle? style) {
        style = null;
        JsonElement colorway = WithTemplateDefault(layout, templateLayout, "colorway");
        if (!TryReadPalette(colorway, out OfficeColor[]? palette) ||
            !TryReadColor(WithTemplateDefault(layout, templateLayout, "paper_bgcolor"), out OfficeColor? paperColor) ||
            !TryReadColor(WithTemplateDefault(layout, templateLayout, "plot_bgcolor"), out OfficeColor? plotColor)) return false;
        style = new OfficeChartStyle(palette: palette, backgroundColor: paperColor,
            plotAreaBackgroundColor: plotColor, showBorder: false);
        return true;
    }

    private static JsonElement WithTemplateDefault(JsonElement layout, JsonElement templateLayout, string name) {
        JsonElement value = Property(layout, name);
        return value.ValueKind == JsonValueKind.Undefined ? Property(templateLayout, name) : value;
    }

    private static bool TryReadPalette(JsonElement source, out OfficeColor[]? palette) {
        palette = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is < 1 or > 32) return false;
        var colors = new OfficeColor[source.GetArrayLength()];
        int index = 0;
        foreach (JsonElement entry in source.EnumerateArray()) {
            if (!TryReadColor(entry, out OfficeColor? color) || !color.HasValue) return false;
            colors[index++] = color.Value;
        }
        palette = colors;
        return true;
    }

    private static bool TryReadColor(JsonElement source, out OfficeColor? color) {
        color = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(source, 128, out string? value) || !OfficeColor.TryParseCss(value, out OfficeColor parsed))
            return false;
        color = parsed;
        return true;
    }

    private readonly struct PlotlyTraceStyle {
        internal PlotlyTraceStyle(OfficeColor? markerColor, OfficeColor? lineColor, OfficeLength? markerSize,
            OfficeLength? lineWidth, OfficeChartMarkerShape? markerShape, OfficeStrokeDashStyle? lineDash) {
            MarkerColor = markerColor;
            LineColor = lineColor;
            MarkerSize = markerSize;
            LineWidth = lineWidth;
            MarkerShape = markerShape;
            LineDash = lineDash;
        }

        internal OfficeColor? MarkerColor { get; }
        internal OfficeColor? LineColor { get; }
        internal OfficeLength? MarkerSize { get; }
        internal OfficeLength? LineWidth { get; }
        internal OfficeChartMarkerShape? MarkerShape { get; }
        internal OfficeStrokeDashStyle? LineDash { get; }
    }
}
