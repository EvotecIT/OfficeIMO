using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

/// <summary>Projects a bounded subset of saved Plotly Cartesian figures to the shared chart renderer.</summary>
internal static partial class NotebookPlotlyChart {
    private const int MaxTraces = 12;
    private const int MaxPoints = 2048;
    private const int MaxWebGlPoints = 8192;
    private const int MaxBarCategories = 64;

    /// <summary>Reads a static grouped bar or numeric scatter figure. Returns null for unsupported or malformed figures.</summary>
    internal static OfficeChartSnapshot? TryRead(JsonElement spec, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!ContainsOnly(spec, "data", "layout", "config")) return null;
        JsonElement traces = Property(spec, "data");
        if (traces.ValueKind != JsonValueKind.Array || traces.GetArrayLength() is < 1 or > MaxTraces) return null;

        JsonElement firstTrace = traces[0];
        JsonElement firstType = Property(firstTrace, "type");
        bool bar = firstType.ValueKind == JsonValueKind.String && firstType.ValueEquals("bar");
        bool webGl = firstType.ValueKind == JsonValueKind.String && firstType.ValueEquals("scattergl");
        bool horizontal = bar && ValueOrDefault(Property(firstTrace, "orientation"), "h", allowMissing: false);
        string traceType = bar ? "bar" : webGl ? "scattergl" : "scatter";
        string? orientation = bar ? (horizontal ? "h" : "v") : null;
        int traceCount = traces.GetArrayLength();

        if (!TryReadLayout(Property(spec, "layout"), traceCount, traceType, horizontal, out PlotlyLayout layout)) return null;

        string[]? categories = null;
        HashSet<string>? offsetGroups = bar && traceCount > 1 ? new HashSet<string>(StringComparer.Ordinal) : null;
        var series = new List<OfficeSeries>(traceCount);
        int remainingPoints = webGl ? MaxWebGlPoints : MaxPoints;
        foreach (JsonElement trace in traces.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (offsetGroups != null && HasProperty(trace, "offsetgroup") &&
                (!TryReadString(Property(trace, "offsetgroup"), MaxLabelLength, out string? group) ||
                    (group!.Length != 0 && !offsetGroups.Add(group)))) return null;
            if (!TryReadTrace(trace, traceType, orientation, layout.ExplicitCategoryAxis, series.Count, remainingPoints,
                    ref categories, out OfficeSeries? result)) return null;
            series.Add(result!);
            remainingPoints -= result!.Values.Count;
        }

        // The shared scatter renderer takes the actual numeric X coordinates from each series.
        categories ??= new[] { "0" };
        var data = new OfficeChartData(categories, series);
        OfficeChartType chartType = bar
            ? (horizontal ? OfficeChartType.Bar : OfficeChartType.Column)
            : OfficeChartType.Scatter;
        return new OfficeChartSnapshot(webGl ? "Plotly WebGL chart" : "Plotly chart", layout.Title, chartType, data, layout.Width, layout.Height,
            layout.Style, layout.Layout);
    }

    private static bool TryReadTrace(JsonElement trace, string traceType, string? orientation, bool explicitCategoryAxis,
        int seriesIndex, int remainingPoints, ref string[]? categories, out OfficeSeries? series,
        string xAxis = "x", string yAxis = "y", bool allowNumericCategories = false) {
        series = null;
        bool bar = orientation != null;
        if (!ContainsOnly(trace, "type", "x", "y", "xaxis", "yaxis", "name", "showlegend", "legendgroup",
                "hoverinfo", "hovertemplate", "hovertext", "customdata", "meta", "marker", "line", "mode",
                "orientation", "alignmentgroup", "offsetgroup", "visible") ||
            !ValueOrDefault(Property(trace, "type"), traceType, allowMissing: traceType == "scatter") ||
            !ValueOrDefault(Property(trace, "xaxis"), xAxis) ||
            !ValueOrDefault(Property(trace, "yaxis"), yAxis) ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showlegend"), true, out bool showLegend) ||
            !TryReadOptionalText(Property(trace, "name"), out string? name)) return false;

        JsonElement x = Property(trace, "x");
        JsonElement y = Property(trace, "y");
        double[]? numericX = null;
        double[]? values;
        if (bar) {
            string expectedOrientation = orientation!;
            bool horizontal = expectedOrientation == "h";
            int maxCategories = Math.Min(MaxBarCategories, remainingPoints);
            JsonElement categoryValues = horizontal ? y : x;
            if (!ValueOrDefault(Property(trace, "orientation"), expectedOrientation, allowMissing: !horizontal) ||
                HasProperty(trace, "mode") || HasProperty(trace, "line") ||
                (!TryReadCategories(categoryValues, maxCategories, explicitCategoryAxis, out string[]? traceCategories) &&
                    !(allowNumericCategories && TryReadNumericCategories(categoryValues, maxCategories, out traceCategories)))) return false;
            if (categories != null && !categories.SequenceEqual(traceCategories!, StringComparer.Ordinal)) return false;
            values = ReadNumbers(horizontal ? x : y, traceCategories!.Length);
            if (values == null || values.Length != traceCategories.Length) return false;
            categories ??= traceCategories;
        } else {
            if (HasProperty(trace, "orientation") || HasProperty(trace, "alignmentgroup") || HasProperty(trace, "offsetgroup")) return false;
            values = ReadNumbers(y, remainingPoints);
            if (values == null) return false;
            numericX = x.ValueKind == JsonValueKind.Undefined ? DefaultScatterX(values.Length) : ReadNumbers(x, values.Length);
            if (numericX == null || numericX.Length != values.Length) return false;
        }

        bool markers = false;
        bool connect = false;
        if (!bar && !TryReadScatterMode(Property(trace, "mode"), values.Length < 20, out markers, out connect)) return false;
        if (!TryReadTraceStyle(trace, bar, markers, connect, out PlotlyTraceStyle style)) return false;
        name ??= "trace " + seriesIndex.ToString(CultureInfo.InvariantCulture);
        series = new OfficeSeries(name, values, numericX, style.MarkerColor ?? style.LineColor, pointColors: null,
            showMarkers: markers, showInLegend: showLegend, connectLine: connect, markerSize: style.MarkerSize,
            markerShape: style.MarkerShape, strokeWidth: style.LineWidth, strokeDashStyle: style.LineDash);
        if (style.LineColor.HasValue && style.MarkerColor.HasValue && style.LineColor != style.MarkerColor) {
            series = series.WithSeriesStyle(new OfficeChartPointStyle(outlineColor: style.LineColor));
        }

        return true;
    }

    private static bool TryReadScatterMode(JsonElement mode, bool defaultMarkers, out bool markers, out bool connect) {
        markers = false;
        connect = false;
        if (mode.ValueKind == JsonValueKind.Undefined) {
            markers = defaultMarkers;
            connect = true;
            return true;
        }
        if (mode.ValueKind != JsonValueKind.String) return false;
        markers = mode.ValueEquals("markers") || mode.ValueEquals("lines+markers") || mode.ValueEquals("markers+lines");
        connect = mode.ValueEquals("lines") || mode.ValueEquals("lines+markers") || mode.ValueEquals("markers+lines");
        return markers || connect;
    }

    private static bool TryReadLayout(JsonElement source, int seriesCount, string traceType, bool horizontal, out PlotlyLayout result) {
        result = default;
        bool bar = traceType == "bar";
        JsonElement xAxis = Property(source, "xaxis");
        JsonElement yAxis = Property(source, "yaxis");
        if (!ContainsOnly(source, "title", "width", "height", "xaxis", "yaxis", "showlegend", "legend", "barmode",
                "scattermode", "template", "margin", "autosize", "hovermode", "hoverlabel", "colorway",
                "paper_bgcolor", "plot_bgcolor") ||
            !TryReadLayoutTitle(Property(source, "title"), out string? title) ||
            !TryReadLength(Property(source, "width"), 200, 1600, out OfficeLength? width) ||
            !TryReadLength(Property(source, "height"), 180, 1200, out OfficeLength? height) ||
            !TryReadAxis(xAxis, bar && !horizontal, "y", out string? xTitle) ||
            !TryReadAxis(yAxis, horizontal, "x", out string? yTitle) ||
            !TryReadBoolean(Property(source, "showlegend"), seriesCount > 1, out bool showLegend) ||
            !TryReadLegend(Property(source, "legend"), out OfficeChartLegendPosition legendPosition) ||
            !ValueOrDefault(Property(source, "scattermode"), "overlay") ||
            !TryReadTemplate(Property(source, "template"), traceType, out JsonElement templateLayout) ||
            !TryReadChartStyle(source, templateLayout, out OfficeChartStyle? style)) return false;

        JsonElement barMode = Property(source, "barmode");
        if (bar && !ValueOrDefault(barMode, "group") &&
            !(seriesCount == 1 && ValueOrDefault(barMode, "relative", allowMissing: false))) return false;

        var layout = new OfficeChartLayout(showLegend: showLegend, legendPosition: legendPosition,
            categoryAxisTitle: horizontal ? yTitle : xTitle, valueAxisTitle: horizontal ? xTitle : yTitle,
            horizontalAxisCrossingPosition: OfficeChartAxisCrossingPosition.Minimum,
            verticalAxisCrossingPosition: OfficeChartAxisCrossingPosition.Minimum);
        bool explicitCategoryAxis = bar && HasProperty(horizontal ? yAxis : xAxis, "type");
        result = new PlotlyLayout(title, width ?? OfficeLength.Pixels(700), height ?? OfficeLength.Pixels(450),
            style!, layout, explicitCategoryAxis);
        return true;
    }

    private static bool TryReadAxis(JsonElement source, bool categorical, string anchor, out string? title) {
        title = null;
        if (!ContainsOnly(source, "title", "type", "anchor", "domain", "automargin") ||
            !ValueOrDefault(Property(source, "anchor"), anchor) ||
            !ValueOrDefault(Property(source, "type"), categorical ? "category" : "linear")) return false;
        JsonElement domain = Property(source, "domain");
        if (domain.ValueKind != JsonValueKind.Undefined &&
            (domain.ValueKind != JsonValueKind.Array || domain.GetArrayLength() != 2 ||
                !TryReadNumber(domain[0], out double start) || start != 0 ||
                !TryReadNumber(domain[1], out double end) || end != 1)) return false;
        return TryReadLayoutTitle(Property(source, "title"), out title);
    }

    private static bool TryReadLegend(JsonElement source, out OfficeChartLegendPosition position) {
        position = OfficeChartLegendPosition.Right;
        if (!ContainsOnly(source, "orientation", "tracegroupgap")) return false;
        JsonElement orientation = Property(source, "orientation");
        if (ValueOrDefault(orientation, "v")) return true;
        if (!ValueOrDefault(orientation, "h", allowMissing: false)) return false;
        position = OfficeChartLegendPosition.Bottom;
        return true;
    }

    private static bool TryReadLayoutTitle(JsonElement element, out string? title) {
        title = null;
        if (element.ValueKind != JsonValueKind.Object) return TryReadOptionalText(element, out title);
        return ContainsOnly(element, "text") && TryReadOptionalText(Property(element, "text"), out title);
    }

    private readonly struct PlotlyLayout {
        internal PlotlyLayout(string? title, OfficeLength width, OfficeLength height, OfficeChartStyle style,
            OfficeChartLayout layout, bool explicitCategoryAxis) {
            Title = title;
            Width = width;
            Height = height;
            Style = style;
            Layout = layout;
            ExplicitCategoryAxis = explicitCategoryAxis;
        }

        internal string? Title { get; }
        internal OfficeLength Width { get; }
        internal OfficeLength Height { get; }
        internal OfficeChartStyle Style { get; }
        internal OfficeChartLayout Layout { get; }
        internal bool ExplicitCategoryAxis { get; }
    }
}
