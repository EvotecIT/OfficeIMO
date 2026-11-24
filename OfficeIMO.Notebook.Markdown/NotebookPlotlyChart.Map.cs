using System.Globalization;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private const int MaximumMapVertices = 16000;
    private const int MaximumMapRegions = 256;

    /// <summary>Renders saved geographic observations and region data without fetching tiles or geometry.</summary>
    internal static OfficeDrawing? TryRenderGeographic(JsonElement spec, CancellationToken cancellationToken, out string? title,
        List<NotebookChartMark>? marks = null) {
        title = null;
        cancellationToken.ThrowIfCancellationRequested();
        JsonElement data = Property(spec, "data");
        if (!ContainsOnly(spec, "data", "layout", "config") || data.ValueKind != JsonValueKind.Array ||
            data.GetArrayLength() is < 1 or > MaxTraces ||
            !TryReadMapLayout(Property(spec, "layout"), data.GetArrayLength(), out MapLayout? layout)) return null;

        var scatter = new List<MapScatter>(data.GetArrayLength());
        MapChoropleth? regions = null;
        int remaining = MaxPoints;
        foreach (JsonElement trace in data.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!ValueOrDefault(Property(trace, "geo"), "geo")) return null;
            JsonElement type = Property(trace, "type");
            if (ValueOrDefault(type, "scattergeo", allowMissing: false)) {
                MapScatter? parsed = ReadMapScatter(trace, layout!.Style.GetSeriesColor(scatter.Count), remaining);
                if (parsed == null) return null;
                remaining -= parsed.Coordinates.Length;
                scatter.Add(parsed);
            } else if (ValueOrDefault(type, "choropleth", allowMissing: false) && regions == null) {
                regions = ReadMapChoropleth(trace, layout!.Template, cancellationToken);
                if (regions == null) return null;
            } else {
                return null;
            }
        }
        OfficeDrawing? drawing = DrawMap(layout!, scatter, regions, cancellationToken, marks);
        if (drawing != null) title = layout!.Title;
        return drawing;
    }

    private static bool TryReadMapLayout(JsonElement source, int traces, out MapLayout? layout) {
        layout = null;
        JsonElement template = Property(source, "template");
        JsonElement geo = Property(source, "geo");
        if (!ContainsOnly(source, "template", "title", "width", "height", "geo", "showlegend", "legend", "paper_bgcolor",
                "colorway", "font", "margin", "hovermode", "hoverlabel", "autosize", "dragmode") ||
            !TryReadLayoutTitle(Property(source, "title"), out string? title) ||
            !TryReadLength(Property(source, "width"), 320, 1600, out OfficeLength? width) ||
            !TryReadLength(Property(source, "height"), 240, 1200, out OfficeLength? height) ||
            !TryReadBoolean(Property(source, "showlegend"), traces > 1, out bool showLegend) ||
            !TryReadLegend(Property(source, "legend"), out OfficeChartLegendPosition legend) ||
            !TryReadTemplate(template, "scatter", out JsonElement templateLayout) ||
            !TryReadChartStyle(source, templateLayout, out OfficeChartStyle? style) ||
            !TryReadSceneFont(WithTemplateDefault(source, templateLayout, "font"), out OfficeColor? textColor) ||
            !TryReadMapSettings(geo, Property(templateLayout, "geo"), out MapSettings? settings)) return false;
        layout = new MapLayout(title, (width ?? OfficeLength.Pixels(700)).ToPoints(), (height ?? OfficeLength.Pixels(450)).ToPoints(),
            style!, textColor ?? style!.TextColor, showLegend, legend, settings!, template);
        return true;
    }

    private static bool TryReadMapSettings(JsonElement geo, JsonElement template, out MapSettings? settings) {
        settings = null;
        JsonElement projection = Property(geo, "projection");
        JsonElement templateProjection = Property(template, "projection");
        if (!ContainsOnly(geo, "projection", "scope", "fitbounds", "lonaxis", "lataxis", "showland", "landcolor", "bgcolor",
                "showcoastlines", "coastlinecolor", "coastlinewidth", "showcountries", "countrycolor", "countrywidth", "showframe",
                "framecolor", "framewidth", "showocean", "oceancolor", "showlakes", "lakecolor", "subunitcolor", "resolution",
                "visible") ||
            !ContainsOnly(projection, "type", "scale") || !ContainsOnly(templateProjection, "type", "scale") ||
            !ValueOrDefault(Property(geo, "scope"), "world") ||
            !TryReadBoolean(Property(geo, "visible"), true, out bool visible) || !visible ||
            !TryReadMapDefaultNumber(WithTemplateDefault(projection, templateProjection, "scale"), 1) ||
            !TryReadMapDefaultNumber(Property(geo, "resolution"), 110) ||
            !TryReadMapAxis(Property(geo, "lonaxis"), 180, out MapAxis longitude) ||
            !TryReadMapAxis(Property(geo, "lataxis"), 90, out MapAxis latitude) ||
            !TryReadMapAppearance(geo, template, out MapAppearance? appearance)) return false;

        JsonElement projectionType = WithTemplateDefault(projection, templateProjection, "type");
        bool mercator = ValueOrDefault(projectionType, "mercator", allowMissing: false);
        if (!mercator && !ValueOrDefault(projectionType, "equirectangular")) return false;

        JsonElement fit = Property(geo, "fitbounds");
        MapFit fitMode;
        if (fit.ValueKind is JsonValueKind.Undefined or JsonValueKind.False) fitMode = MapFit.World;
        else if (ValueOrDefault(fit, "locations", allowMissing: false)) fitMode = MapFit.Locations;
        else if (ValueOrDefault(fit, "geojson", allowMissing: false)) fitMode = MapFit.GeoJson;
        else return false;
        settings = new MapSettings(mercator, fitMode, longitude, latitude, appearance!);
        return true;
    }

    private static bool TryReadMapDefaultNumber(JsonElement value, double expected) =>
        value.ValueKind == JsonValueKind.Undefined || TryReadNumber(value, out double parsed) && parsed == expected;

    private static bool TryReadMapAxis(JsonElement source, double range, out MapAxis result) {
        result = default;
        if (!ContainsOnly(source, "range", "showgrid", "gridcolor", "gridwidth", "dtick", "tick0") ||
            !TryReadBoolean(Property(source, "showgrid"), false, out bool showGrid) ||
            !TryReadColor(Property(source, "gridcolor"), out OfficeColor? gridColor) ||
            !TryReadLength(Property(source, "gridwidth"), 0, 8, out OfficeLength? gridWidth) ||
            !TryReadMapAxisNumber(Property(source, "tick0"), 0, -range, range, out double gridOrigin) ||
            !TryReadMapAxisNumber(Property(source, "dtick"), 0, 0.001, range, out double gridStep)) return false;
        double[]? bounds = ReadNumbers(Property(source, "range"), 2);
        if (HasProperty(source, "range") && (bounds?.Length != 2 || bounds[0] < -range || bounds[1] > range || bounds[0] >= bounds[1]))
            return false;
        result = new MapAxis(bounds, showGrid, gridColor ?? OfficeColor.FromRgb(202, 212, 220),
            gridWidth?.ToPoints() ?? 0.45, gridOrigin, gridStep);
        return true;
    }

    private static bool TryReadMapAxisNumber(JsonElement source, double fallback, double min, double max, out double number) {
        number = fallback;
        return source.ValueKind == JsonValueKind.Undefined || TryReadNumber(source, out number) && number >= min && number <= max;
    }

    private static bool TryReadMapAppearance(JsonElement geo, JsonElement template, out MapAppearance? appearance) {
        appearance = null;
        if (!TryReadBoolean(WithTemplateDefault(geo, template, "showland"), true, out bool showLand) ||
            !TryReadBoolean(WithTemplateDefault(geo, template, "showcoastlines"), true, out bool showCoast) ||
            !TryReadBoolean(WithTemplateDefault(geo, template, "showcountries"), false, out bool showCountries) ||
            !TryReadBoolean(WithTemplateDefault(geo, template, "showframe"), true, out bool showFrame) ||
            !TryReadBoolean(WithTemplateDefault(geo, template, "showocean"), false, out bool showOcean) ||
            !TryReadBoolean(Property(geo, "showlakes"), false, out bool showLakes) || showLakes ||
            (HasProperty(geo, "lakecolor") && Property(geo, "showlakes").ValueKind != JsonValueKind.False &&
                Property(template, "showlakes").ValueKind == JsonValueKind.True) ||
            !TryReadColor(WithTemplateDefault(geo, template, "landcolor"), out OfficeColor? land) ||
            !TryReadColor(WithTemplateDefault(geo, template, "coastlinecolor"), out OfficeColor? coast) ||
            !TryReadColor(WithTemplateDefault(geo, template, "countrycolor"), out OfficeColor? borders) ||
            !TryReadColor(WithTemplateDefault(geo, template, "bgcolor"), out OfficeColor? background) ||
            !TryReadColor(WithTemplateDefault(geo, template, "oceancolor"), out OfficeColor? ocean) ||
            !TryReadColor(WithTemplateDefault(geo, template, "framecolor"), out OfficeColor? frame) ||
            !TryReadLength(WithTemplateDefault(geo, template, "coastlinewidth"), 0, 8, out OfficeLength? coastWidth) ||
            !TryReadLength(WithTemplateDefault(geo, template, "countrywidth"), 0, 8, out OfficeLength? borderWidth) ||
            !TryReadLength(WithTemplateDefault(geo, template, "framewidth"), 0, 8, out OfficeLength? frameWidth)) return false;
        appearance = new MapAppearance(
            showOcean ? ocean ?? OfficeColor.FromRgb(220, 237, 248) : background ?? OfficeColor.White,
            showLand ? land ?? OfficeColor.FromRgb(225, 231, 236) : null,
            showCoast ? coast ?? OfficeColor.FromRgb(105, 122, 135) : null, coastWidth?.ToPoints() ?? 0.65,
            showCountries ? borders ?? OfficeColor.FromRgb(151, 161, 170) : null, borderWidth?.ToPoints() ?? 0.5,
            showFrame ? frame ?? OfficeColor.FromRgb(116, 130, 139) : null, frameWidth?.ToPoints() ?? 0.65);
        return true;
    }

    private static MapScatter? ReadMapScatter(JsonElement trace, OfficeColor seriesColor, int remaining) {
        if (!ContainsOnly(trace, "type", "geo", "lat", "lon", "locations", "locationmode", "name", "mode", "marker", "line",
                "showlegend", "legendgroup", "hoverinfo", "hovertext", "hovertemplate", "text", "textposition", "textfont",
                "customdata", "meta", "opacity", "connectgaps", "visible") ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showlegend"), true, out bool legend) ||
            !TryReadBoolean(Property(trace, "connectgaps"), false, out bool connectGaps) ||
            !TryReadUnitInterval(Property(trace, "opacity"), 1, out double opacity) ||
            !TryReadOptionalText(Property(trace, "name"), out string? name)) return null;
        MapCoordinate?[]? positions = ReadMapCoordinates(trace, remaining);
        if (positions == null || !TryReadMapMode(Property(trace, "mode"), positions.Length, out bool markers, out bool lines, out bool text) ||
            !TryReadMapScatterStyle(trace, seriesColor, opacity, out MapScatterStyle style)) return null;
        string?[]? labels = ReadMapLabels(Property(trace, "text"), positions.Length);
        string?[]? hover = ReadMapLabels(Property(trace, "hovertext"), positions.Length);
        if ((HasProperty(trace, "text") && labels == null) || (HasProperty(trace, "hovertext") && hover == null)) return null;
        return new MapScatter(positions, markers, lines, text, connectGaps, style, text ? labels : null, hover ?? labels, name, legend);
    }

    private static bool TryReadMapMode(JsonElement value, int count, out bool markers, out bool lines, out bool text) {
        text = false;
        if (TryReadScatterMode(value, count < 20, out markers, out lines)) return true;
        if (!TryReadString(value, 18, out string? mode)) return false;
        foreach (string part in mode!.Split('+')) {
            if (part == "text" && !text) text = true;
            else if (part == "markers" && !markers) markers = true;
            else if (part == "lines" && !lines) lines = true;
            else return false;
        }
        return text;
    }

    private static bool TryReadMapScatterStyle(JsonElement trace, OfficeColor color, double opacity, out MapScatterStyle style) {
        style = default;
        JsonElement marker = Property(trace, "marker");
        JsonElement outline = Property(marker, "line");
        JsonElement line = Property(trace, "line");
        if (!ContainsOnly(marker, "color", "size", "symbol", "line", "opacity") ||
            !ContainsOnly(outline, "color", "width") ||
            !ContainsOnly(line, "color", "width", "dash") ||
            !TryReadColor(Property(marker, "color"), out OfficeColor? markerColor) ||
            !TryReadLength(Property(marker, "size"), 0.5, 64, out OfficeLength? size) ||
            !TryReadUnitInterval(Property(marker, "opacity"), 1, out double markerOpacity) ||
            !TryReadMarkerShape(Property(marker, "symbol"), out OfficeChartMarkerShape? symbol) ||
            symbol == OfficeChartMarkerShape.Triangle ||
            !TryReadColor(Property(outline, "color"), out OfficeColor? outlineColor) ||
            !TryReadLength(Property(outline, "width"), 0, 12, out OfficeLength? outlineWidth) ||
            !TryReadColor(Property(line, "color"), out OfficeColor? lineColor) ||
            !TryReadLength(Property(line, "width"), 0.1, 24, out OfficeLength? lineWidth) ||
            !TryReadLineDash(Property(line, "dash"), out OfficeStrokeDashStyle? lineDash)) return false;
        style = new MapScatterStyle(WithMapOpacity(markerColor ?? color, opacity * markerOpacity),
            WithMapOpacity(outlineColor ?? OfficeColor.White, opacity * markerOpacity), WithMapOpacity(lineColor ?? color, opacity),
            size?.ToPoints() ?? 5.5, outlineWidth?.ToPoints() ?? 0, symbol ?? OfficeChartMarkerShape.Circle,
            lineWidth?.ToPoints() ?? 1.7, lineDash ?? OfficeStrokeDashStyle.Solid);
        return true;
    }

    private static OfficeColor WithMapOpacity(OfficeColor color, double opacity) =>
        OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * opacity));

    private sealed record MapLayout(string? Title, double Width, double Height, OfficeChartStyle Style,
        OfficeColor TextColor, bool ShowLegend, OfficeChartLegendPosition Legend, MapSettings Settings, JsonElement Template);

    private sealed record MapSettings(bool Mercator, MapFit Fit, MapAxis Longitude, MapAxis Latitude, MapAppearance Appearance);

    private readonly record struct MapAxis(double[]? Range, bool ShowGrid, OfficeColor GridColor, double GridWidth, double Origin, double Step);

    private enum MapFit { World, Locations, GeoJson }

    private sealed record MapAppearance(OfficeColor Background, OfficeColor? Land, OfficeColor? Coast, double CoastWidth,
        OfficeColor? Borders, double BorderWidth, OfficeColor? Frame, double FrameWidth);

    private sealed record MapScatter(MapCoordinate?[] Coordinates, bool Markers, bool Lines, bool Text, bool ConnectGaps,
        MapScatterStyle Style, string?[]? Labels, string?[]? Hover, string? Name, bool Legend);

    private readonly record struct MapScatterStyle(OfficeColor MarkerColor, OfficeColor OutlineColor, OfficeColor LineColor,
        double MarkerSize, double OutlineWidth, OfficeChartMarkerShape Shape, double LineWidth, OfficeStrokeDashStyle Dash);

    private readonly record struct MapCoordinate(double Longitude, double Latitude);
}
