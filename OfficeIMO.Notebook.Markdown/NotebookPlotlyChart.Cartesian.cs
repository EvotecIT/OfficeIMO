using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaxCartesianPanels = 8;

    /// <summary>Draws saved heatmaps and independent Cartesian panels from explicit Plotly axis domains.</summary>
    internal static OfficeDrawing? TryRenderCartesianPanels(JsonElement spec, CancellationToken cancellationToken,
        out string? title, List<NotebookChartMark>? marks = null) {
        title = null;
        cancellationToken.ThrowIfCancellationRequested();
        JsonElement traces = Property(spec, "data");
        JsonElement layout = Property(spec, "layout");
        if (traces.ValueKind != JsonValueKind.Array || traces.GetArrayLength() is < 1 or > MaxTraces ||
            !ContainsOnly(spec, "data", "layout", "config") ||
            !TryReadCartesianPanels(traces, layout, out List<CartesianPanel> panels) ||
            (panels.Count == 1 && !panels[0].IsHeatmap) ||
            !TryReadCartesianLayout(layout, panels, out CartesianLayout appearance)) return null;

        var drawing = new OfficeDrawing(appearance.Width, appearance.Height);
        PaintCartesianRectangle(drawing, new CartesianBox(0, 0, drawing.Width, drawing.Height), appearance.Style.BackgroundColor);
        if (!string.IsNullOrWhiteSpace(appearance.Title)) {
            PaintSceneLabel(drawing, appearance.Title!, drawing.Width / 2, 17, drawing.Width - 24, 15, appearance.Style.TitleColor);
        }

        var heatmapColorAxes = new HashSet<string>(StringComparer.Ordinal);
        List<CartesianLegendItem>? legend = appearance.LegendBox.HasValue ? new List<CartesianLegendItem>() : null;
        int remainingPoints = MaxWebGlPoints;
        int remainingHeatCells = MaxHeatCells;
        int firstMark = marks?.Count ?? 0;
        foreach (CartesianPanel panel in panels) {
            cancellationToken.ThrowIfCancellationRequested();
            CartesianBox box = GetCartesianPanelBox(appearance.Content, panel);
            bool rendered = panel.IsHeatmap
                ? TryPaintHeatmap(drawing, panel, box, appearance, heatmapColorAxes, ref remainingHeatCells, cancellationToken, marks)
                : TryPaintCartesianSeries(drawing, panel, box, appearance, ref remainingPoints, cancellationToken, marks, legend);
            if (!rendered) {
                if (marks != null) marks.RemoveRange(firstMark, marks.Count - firstMark);
                return null;
            }
        }
        if (legend is { Count: > 0 } && appearance.LegendBox is CartesianBox legendBox) {
            PaintCartesianLegend(drawing, legendBox, appearance.LegendPosition, legend, appearance.Style.TextColor);
        }
        title = appearance.Title;
        return drawing;
    }

    private static bool TryReadCartesianPanels(JsonElement traces, JsonElement layout, out List<CartesianPanel> panels) {
        panels = new List<CartesianPanel>();
        var byAxes = new Dictionary<(string X, string Y), CartesianPanel>();
        int traceIndex = 0;
        foreach (JsonElement trace in traces.EnumerateArray()) {
            if (trace.ValueKind != JsonValueKind.Object ||
                !TryReadCartesianId(Property(trace, "xaxis"), "x", MaxCartesianPanels, out string? x) ||
                !TryReadCartesianId(Property(trace, "yaxis"), "y", MaxCartesianPanels, out string? y) ||
                !TryReadCartesianTraceType(trace, out string kind)) return false;

            if (byAxes.TryGetValue((x!, y!), out CartesianPanel? existing)) {
                if (existing.Kind != kind || existing.IsHeatmap) return false;
                existing.Traces.Add(new CartesianTrace(trace, traceIndex++));
                continue;
            }

            if (panels.Count == MaxCartesianPanels ||
                !TryReadCartesianAxis(Property(layout, CartesianAxisProperty(x!)), y!, out CartesianAxis xAxis) ||
                !TryReadCartesianAxis(Property(layout, CartesianAxisProperty(y!)), x!, out CartesianAxis yAxis) ||
                panels.Any(panel => panel.XId == x || panel.YId == y ||
                    (DomainsOverlap(panel.X, xAxis) && DomainsOverlap(panel.Y, yAxis)))) return false;
            var nextPanel = new CartesianPanel(x!, y!, kind, xAxis, yAxis);
            nextPanel.Traces.Add(new CartesianTrace(trace, traceIndex++));
            panels.Add(nextPanel);
            byAxes.Add((x!, y!), nextPanel);
        }
        return true;
    }

    private static bool TryReadCartesianLayout(JsonElement layout, List<CartesianPanel> panels, out CartesianLayout result) {
        result = default;
        if (layout.ValueKind != JsonValueKind.Undefined && layout.ValueKind != JsonValueKind.Object) return false;
        JsonElement template = Property(layout, "template");
        if (layout.ValueKind == JsonValueKind.Object && layout.EnumerateObject().Any(property =>
            !IsCartesianLayoutProperty(property.Name, panels))) return false;
        if (!TryReadLayoutTitle(Property(layout, "title"), out string? title) ||
            !TryReadLength(Property(layout, "width"), 200, 1600, out OfficeLength? width) ||
            !TryReadLength(Property(layout, "height"), 180, 1200, out OfficeLength? height) ||
            !TryReadBoolean(Property(layout, "showlegend"), panels.Sum(panel => panel.Traces.Count) > 1, out bool legend) ||
            !TryReadLegend(Property(layout, "legend"), out OfficeChartLegendPosition legendPosition) ||
            !ValueOrDefault(Property(layout, "scattermode"), "overlay") ||
            !TryReadPanelTitles(Property(layout, "annotations"), panels) ||
            !TryReadCartesianMargins(Property(layout, "margin"), title, out CartesianMargins margins)) return false;

        JsonElement templateLayout = Property(template, "layout");
        // Check defaults for each active trace family, including defaults inherited by numbered axes.
        foreach (string kind in panels.Select(panel => panel.Kind).Distinct(StringComparer.Ordinal)) {
            if (!TryReadTemplate(template, kind, out _)) return false;
        }
        if (!TryReadChartStyle(layout, templateLayout, out OfficeChartStyle? style)) return false;
        double chartWidth = (width ?? OfficeLength.Pixels(700)).ToPoints();
        double chartHeight = (height ?? OfficeLength.Pixels(450)).ToPoints();
        var content = new CartesianBox(margins.Left, margins.Top,
            chartWidth - margins.Left - margins.Right, chartHeight - margins.Top - margins.Bottom);
        if (content.Width < 120 || content.Height < 100) return false;
        int legendCount = legend ? panels.Where(panel => !panel.IsHeatmap).SelectMany(panel => panel.Traces)
            .Count(trace => TryReadBoolean(Property(trace.Source, "showlegend"), true, out bool visible) && visible) : 0;
        CartesianBox? legendBox = ReserveCartesianLegend(legendCount, legendPosition, ref content);
        result = new CartesianLayout(chartWidth, chartHeight, title, style!, content, layout, template, legendBox, legendPosition);
        return true;
    }

    private static bool IsCartesianLayoutProperty(string key, List<CartesianPanel> panels) => key switch {
        "title" or "width" or "height" or "showlegend" or "legend" or "barmode" or "scattermode" or
        "template" or "margin" or "autosize" or "hovermode" or "hoverlabel" or "colorway" or
        "paper_bgcolor" or "plot_bgcolor" or "annotations" => true,
        _ => panels.Any(panel => key == CartesianAxisProperty(panel.XId) || key == CartesianAxisProperty(panel.YId)) ||
            IsCartesianId(key, "coloraxis", MaxCartesianPanels)
    };

    private static bool TryReadCartesianTraceType(JsonElement trace, out string kind) {
        kind = "scatter";
        JsonElement type = Property(trace, "type");
        if (type.ValueKind == JsonValueKind.Undefined) return true;
        if (!TryReadString(type, 16, out string? actual)) return false;
        kind = actual!;
        return kind is "scatter" or "scattergl" or "bar" or "heatmap";
    }

    private static bool TryReadCartesianId(JsonElement input, string prefix, int limit, out string? id) {
        id = prefix;
        if (input.ValueKind == JsonValueKind.Undefined) return true;
        return TryReadString(input, prefix.Length + 1, out id) && IsCartesianId(id!, prefix, limit);
    }

    private static bool IsCartesianId(string value, string prefix, int limit) =>
        value == prefix || (value.Length == prefix.Length + 1 && value.StartsWith(prefix, StringComparison.Ordinal) &&
            value[prefix.Length] >= '2' && value[prefix.Length] <= '0' + limit);

    private static string CartesianAxisProperty(string id) => id.Substring(0, 1) + "axis" + id.Substring(1);

    private static bool TryReadCartesianAxis(JsonElement source, string anchor, out CartesianAxis axis) {
        axis = default;
        if (!ContainsOnly(source, "title", "type", "anchor", "domain", "automargin", "showticklabels",
                "range", "autorange", "scaleanchor", "constrain") ||
            !ValueOrDefault(Property(source, "anchor"), anchor) ||
            !TryReadLayoutTitle(Property(source, "title"), out string? title) ||
            !TryReadBoolean(Property(source, "showticklabels"), true, out bool showLabels) ||
            !TryReadPair(Property(source, "domain"), out double? start, out double? end, 0, 1) ||
            (start.HasValue && start >= end) ||
            !TryReadPair(Property(source, "range"), out double? min, out double? max, -MaxNumberMagnitude, MaxNumberMagnitude) ||
            (min.HasValue && min == max)) return false;

        JsonElement type = Property(source, "type");
        bool categorical = ValueOrDefault(type, "category", allowMissing: false);
        bool linear = ValueOrDefault(type, "linear", allowMissing: false);
        if ((!categorical && !linear && type.ValueKind != JsonValueKind.Undefined) || (categorical && min.HasValue))
            return false;
        JsonElement autorange = Property(source, "autorange");
        bool reversed = ValueOrDefault(autorange, "reversed", allowMissing: false);
        if (!reversed && !TryReadBoolean(autorange, true, out _)) return false;
        if (min.HasValue && max < min) {
            (min, max) = (max, min);
            reversed = true;
        }
        bool lockAspect = HasProperty(source, "scaleanchor");
        if (!ValueOrDefault(Property(source, "scaleanchor"), anchor) ||
            !ValueOrDefault(Property(source, "constrain"), "domain")) return false;
        axis = new CartesianAxis(start ?? 0, end ?? 1, min, max, title, showLabels, categorical, linear, reversed, lockAspect);
        return true;
    }

    private static bool TryReadPair(JsonElement source, out double? first, out double? second, double min, double max) {
        first = second = null;
        if (source.ValueKind == JsonValueKind.Undefined) return true;
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() != 2 ||
            !TryReadNumber(source[0], out double firstValue) || !TryReadNumber(source[1], out double secondValue) ||
            firstValue < min || firstValue > max || secondValue < min || secondValue > max) return false;
        first = firstValue;
        second = secondValue;
        return true;
    }

    private static bool TryReadPanelTitles(JsonElement annotations, List<CartesianPanel> panels) {
        if (annotations.ValueKind == JsonValueKind.Undefined) return true;
        if (annotations.ValueKind != JsonValueKind.Array || annotations.GetArrayLength() > panels.Count) return false;
        foreach (JsonElement annotation in annotations.EnumerateArray()) {
            if (!ContainsOnly(annotation, "text", "x", "y", "xref", "yref", "xanchor", "yanchor", "showarrow", "font") ||
                !ContainsOnly(Property(annotation, "font"), "size", "color", "family") ||
                !TryReadOptionalText(Property(annotation, "text"), out string? title) || string.IsNullOrWhiteSpace(title) ||
                !TryReadNumber(Property(annotation, "x"), out double x) ||
                !TryReadNumber(Property(annotation, "y"), out double y) ||
                !ValueOrDefault(Property(annotation, "xref"), "paper", allowMissing: false) ||
                !ValueOrDefault(Property(annotation, "yref"), "paper", allowMissing: false) ||
                !ValueOrDefault(Property(annotation, "xanchor"), "center") ||
                !ValueOrDefault(Property(annotation, "yanchor"), "bottom") ||
                !TryReadBoolean(Property(annotation, "showarrow"), false, out bool showArrow) || showArrow) return false;
            CartesianPanel? panel = panels.FirstOrDefault(candidate =>
                Math.Abs((candidate.X.Start + candidate.X.End) / 2 - x) < 1e-8 && Math.Abs(candidate.Y.End - y) < 1e-8);
            if (panel == null || panel.Title != null) return false;
            panel.Title = title;
        }
        return true;
    }

    private static bool TryReadCartesianMargins(JsonElement source, string? title, out CartesianMargins margins) {
        margins = default;
        if (!ContainsOnly(source, "l", "r", "t", "b", "pad", "autoexpand") ||
            !TryReadLength(Property(source, "l"), 0, 600, out OfficeLength? left) ||
            !TryReadLength(Property(source, "r"), 0, 600, out OfficeLength? right) ||
            !TryReadLength(Property(source, "t"), 0, 400, out OfficeLength? top) ||
            !TryReadLength(Property(source, "b"), 0, 400, out OfficeLength? bottom)) return false;
        margins = new CartesianMargins(left?.ToPoints() ?? 16, Math.Max(1, right?.ToPoints() ?? 16),
            Math.Max(top?.ToPoints() ?? 10, string.IsNullOrWhiteSpace(title) ? 0 : 32), Math.Max(1, bottom?.ToPoints() ?? 12));
        return true;
    }

    private static CartesianBox GetCartesianPanelBox(CartesianBox content, CartesianPanel panel) =>
        new(content.Left + content.Width * panel.X.Start, content.Top + content.Height * (1 - panel.Y.End),
            content.Width * (panel.X.End - panel.X.Start), content.Height * (panel.Y.End - panel.Y.Start));

    private static bool DomainsOverlap(CartesianAxis first, CartesianAxis second) =>
        Math.Min(first.End, second.End) - Math.Max(first.Start, second.Start) > 1e-8;

    private static void PaintCartesianRectangle(OfficeDrawing drawing, CartesianBox box, OfficeColor color) {
        OfficeShape fill = OfficeShape.Rectangle(box.Width, box.Height);
        fill.FillColor = color;
        drawing.AddShape(fill, box.Left, box.Top);
    }

    private sealed class CartesianPanel {
        internal CartesianPanel(string xId, string yId, string kind, CartesianAxis x, CartesianAxis y) {
            XId = xId;
            YId = yId;
            Kind = kind;
            X = x;
            Y = y;
        }

        internal string XId { get; }
        internal string YId { get; }
        internal string Kind { get; }
        internal CartesianAxis X { get; }
        internal CartesianAxis Y { get; }
        internal List<CartesianTrace> Traces { get; } = new();
        internal string? Title { get; set; }
        internal bool IsHeatmap => Kind == "heatmap";
    }

    private readonly record struct CartesianTrace(JsonElement Source, int Index);
    private readonly record struct CartesianAxis(double Start, double End, double? Minimum, double? Maximum, string? Title,
        bool ShowLabels, bool Categorical, bool Linear, bool Reversed, bool LockAspect);
    private readonly record struct CartesianMargins(double Left, double Right, double Top, double Bottom);
    private readonly record struct CartesianLayout(double Width, double Height, string? Title, OfficeChartStyle Style, CartesianBox Content,
        JsonElement Source, JsonElement Template, CartesianBox? LegendBox, OfficeChartLegendPosition LegendPosition);
    private readonly record struct CartesianBox(double Left, double Top, double Width, double Height) {
        internal double Right => Left + Width;
        internal double Bottom => Top + Height;
    }
}
