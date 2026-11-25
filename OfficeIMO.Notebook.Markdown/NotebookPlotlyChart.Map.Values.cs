using System.Globalization;
using System.IO.Compression;
using System.Text.Json;

namespace OfficeIMO.Notebook.Markdown;

internal static partial class NotebookPlotlyChart {
    private static readonly Lazy<MapWorld> GeographicWorld = new(LoadMapWorld);

    private static MapCoordinate?[]? ReadMapCoordinates(JsonElement trace, int maximum) {
        if (maximum <= 0) return null;
        JsonElement locations = Property(trace, "locations");
        JsonElement longitude = Property(trace, "lon");
        JsonElement latitude = Property(trace, "lat");
        if (locations.ValueKind != JsonValueKind.Undefined) {
            if (longitude.ValueKind != JsonValueKind.Undefined || latitude.ValueKind != JsonValueKind.Undefined ||
                !ValueOrDefault(Property(trace, "locationmode"), "ISO-3") ||
                locations.ValueKind != JsonValueKind.Array || locations.GetArrayLength() is 0 ||
                locations.GetArrayLength() > maximum) return null;
            var located = new MapCoordinate?[locations.GetArrayLength()];
            int index = 0;
            foreach (JsonElement location in locations.EnumerateArray()) {
                if (!TryReadString(location, 3, out string? code) ||
                    !GeographicWorld.Value.CountriesByCode.TryGetValue(code!, out MapFeature? country)) return null;
                located[index++] = country.LabelPosition;
            }
            return located;
        }

        double[]? x = ReadMapPositionValues(longitude, maximum);
        if (x == null) return null;
        double[]? y = ReadMapPositionValues(latitude, x.Length);
        if (y?.Length != x.Length) return null;
        var positions = new MapCoordinate?[x.Length];
        bool hasPosition = false;
        for (int index = 0; index < positions.Length; index++) {
            if (double.IsNaN(x[index]) || double.IsNaN(y[index])) continue;
            if (!IsMapPosition(x[index], y[index])) return null;
            hasPosition = true;
            positions[index] = new MapCoordinate(x[index], y[index]);
        }
        return hasPosition ? positions : null;
    }

    private static double[]? ReadMapPositionValues(JsonElement source, int maximum) {
        if (source.ValueKind == JsonValueKind.Object) return ReadTypedNumbers(source, maximum, allowNaN: true);
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is 0 || source.GetArrayLength() > maximum) return null;
        var values = new double[source.GetArrayLength()];
        int index = 0;
        foreach (JsonElement value in source.EnumerateArray()) {
            if (value.ValueKind == JsonValueKind.Null) values[index] = double.NaN;
            else if (!TryReadNumber(value, out values[index])) return null;
            index++;
        }
        return values;
    }

    private static string?[]? ReadMapLabels(JsonElement source, int count) {
        if (source.ValueKind == JsonValueKind.Undefined) return null;
        var labels = new string?[count];
        if (source.ValueKind == JsonValueKind.String) {
            if (!TryReadOptionalText(source, out string? label)) return null;
            for (int i = 0; i < count; i++) labels[i] = label;
            return labels;
        }
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() != count) return null;
        int at = 0;
        foreach (JsonElement value in source.EnumerateArray()) {
            if (value.ValueKind != JsonValueKind.Null && !TryReadOptionalText(value, out labels[at])) return null;
            at++;
        }
        return labels;
    }

    private static MapChoropleth? ReadMapChoropleth(JsonElement trace, JsonElement template, CancellationToken cancellationToken) {
        JsonElement geoJson = Property(trace, "geojson");
        JsonElement featureKey = Property(trace, "featureidkey");
        JsonElement locationMode = Property(trace, "locationmode");
        JsonElement locations = Property(trace, "locations");
        JsonElement marker = Property(trace, "marker");
        JsonElement line = Property(marker, "line");
        JsonElement colorbar = Property(trace, "colorbar");
        if (!ContainsOnly(trace, "type", "geo", "geojson", "featureidkey", "locations", "locationmode", "z", "zmin", "zmax",
                "colorscale", "reversescale", "showscale", "colorbar", "marker", "name", "text", "hovertext", "hoverinfo",
                "hovertemplate", "customdata", "meta", "visible", "showlegend", "legendgroup", "autocolorscale", "zauto") ||
            locations.ValueKind != JsonValueKind.Array || locations.GetArrayLength() is < 1 or > MaximumMapRegions ||
            !ContainsOnly(marker, "line", "opacity") || !ContainsOnly(line, "color", "width") ||
            !TryReadColor(Property(line, "color"), out OfficeColor? border) ||
            !TryReadLength(Property(line, "width"), 0, 12, out OfficeLength? borderWidth) ||
            !TryReadUnitInterval(Property(marker, "opacity"), 1, out double opacity) ||
            !TryReadBoolean(Property(trace, "visible"), true, out bool visible) || !visible ||
            !TryReadBoolean(Property(trace, "showscale"), true, out bool showScale) ||
            !TryReadBoolean(Property(trace, "reversescale"), false, out bool reverse) ||
            !TryReadLayoutTitle(Property(colorbar, "title"), out string? scaleTitle)) return null;

        int count = locations.GetArrayLength();
        double[]? values = ReadNumbers(Property(trace, "z"), count);
        if (values?.Length != count || !TryReadMapScale(trace, template, values, showScale, reverse, out SceneScale? scale)) return null;

        IReadOnlyDictionary<string, MapFeature> features;
        MapBounds? allBounds = null;
        if (geoJson.ValueKind == JsonValueKind.Undefined) {
            if (!ValueOrDefault(locationMode, "ISO-3") || featureKey.ValueKind != JsonValueKind.Undefined) return null;
            features = GeographicWorld.Value.CountriesByCode;
        } else {
            if (!ValueOrDefault(locationMode, "geojson-id") && !ValueOrDefault(locationMode, "ISO-3")) return null;
            if (!TryReadMapFeatureCollection(geoJson, featureKey, cancellationToken, out var inline, out MapBounds bounds)) return null;
            features = inline!;
            allBounds = bounds;
        }

        string?[]? labels = ReadMapLabels(Property(trace, "hovertext"), count) ?? ReadMapLabels(Property(trace, "text"), count);
        var regions = new List<MapRegion>(count);
        var usedKeys = new HashSet<string>(StringComparer.Ordinal);
        int index = 0;
        foreach (JsonElement location in locations.EnumerateArray()) {
            if (!TryReadMapFeatureKey(location, out string? key) || !features.TryGetValue(key!, out MapFeature? feature)
                || !usedKeys.Add(feature.Key)) return null;
            regions.Add(new MapRegion(feature, values[index], labels?[index] ?? feature.Name));
            index++;
        }
        return new MapChoropleth(regions, allBounds, scale!, scaleTitle, border ?? OfficeColor.White, borderWidth?.ToPoints() ?? 0.7, opacity);
    }

    private static bool TryReadMapScale(JsonElement trace, JsonElement template, double[] values, bool show, bool reverse,
        out SceneScale? scale) {
        scale = null;
        double min = values.Min();
        double max = values.Max();
        if (!TryReadMapAxisNumber(Property(trace, "zmin"), min, -MaxNumberMagnitude, MaxNumberMagnitude, out min) ||
            !TryReadMapAxisNumber(Property(trace, "zmax"), max, -MaxNumberMagnitude, MaxNumberMagnitude, out max) || min > max)
            return false;
        JsonElement source = Property(trace, "colorscale");
        if (source.ValueKind == JsonValueKind.Undefined) {
            JsonElement defaults = Property(Property(template, "data"), "choropleth");
            if (defaults.ValueKind == JsonValueKind.Array && defaults.GetArrayLength() == 1)
                source = Property(defaults[0], "colorscale");
        }
        if (source.ValueKind == JsonValueKind.Undefined)
            source = Property(Property(Property(template, "layout"), "colorscale"), "sequential");
        SceneColorStop[]? stops = ReadSceneColorStops(source);
        if (stops == null) return false;
        scale = new SceneScale(stops, min, max, reverse, show);
        return true;
    }

    private static bool TryReadMapFeatureCollection(JsonElement source, JsonElement matchPath, CancellationToken cancellationToken,
        out Dictionary<string, MapFeature>? index, out MapBounds bounds) {
        index = null;
        bounds = MapBounds.Empty;
        string[] path = new[] { "id" };
        if (matchPath.ValueKind != JsonValueKind.Undefined) {
            if (!TryReadString(matchPath, 90, out string? selector)) return false;
            path = selector!.Split('.');
            if (path.Length > 6 || path.Any(string.IsNullOrWhiteSpace)) return false;
        }
        JsonElement type = Property(source, "type");
        JsonElement features;
        if (ValueOrDefault(type, "FeatureCollection", allowMissing: false)) {
            features = Property(source, "features");
            if (features.ValueKind != JsonValueKind.Array || features.GetArrayLength() is < 1 or > MaximumMapRegions) return false;
        } else if (ValueOrDefault(type, "Feature", allowMissing: false)) {
            features = default;
        } else {
            // An external URL or unknown geometry never participates in document rendering.
            return false;
        }

        var result = new Dictionary<string, MapFeature>(StringComparer.Ordinal);
        int budget = MaximumMapVertices;
        if (features.ValueKind == JsonValueKind.Undefined) {
            if (!TryAddMapFeature(source, path, result, ref budget, ref bounds)) return false;
        } else {
            foreach (JsonElement feature in features.EnumerateArray()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!TryAddMapFeature(feature, path, result, ref budget, ref bounds)) return false;
            }
        }
        index = result;
        return true;
    }

    private static bool TryAddMapFeature(JsonElement source, string[] path, Dictionary<string, MapFeature> index,
        ref int budget, ref MapBounds bounds) {
        if (!ValueOrDefault(Property(source, "type"), "Feature", allowMissing: false)) return false;
        JsonElement id = source;
        foreach (string step in path) id = Property(id, step);
        if (!TryReadMapFeatureKey(id, out string? key) || index.ContainsKey(key!)) return false;
        JsonElement geometry = Property(source, "geometry");
        JsonElement type = Property(geometry, "type");
        JsonElement coordinates = Property(geometry, "coordinates");
        MapPolygon[]? polygons;
        if (ValueOrDefault(type, "Polygon", allowMissing: false)) {
            MapPolygon? polygon = ReadMapPolygon(coordinates, ref budget, requireAntimeridianSplit: true);
            polygons = polygon == null ? null : new[] { polygon };
        } else if (ValueOrDefault(type, "MultiPolygon", allowMissing: false)) {
            polygons = ReadMapPolygons(coordinates, ref budget, requireAntimeridianSplit: true);
        } else {
            return false;
        }
        if (polygons == null) return false;
        MapBounds featureBounds = GetMapBounds(polygons);
        TryReadOptionalText(Property(Property(source, "properties"), "name"), out string? name);
        index.Add(key!, new MapFeature(key!, name ?? key!, featureBounds.Center, polygons, featureBounds));
        bounds = bounds.Include(featureBounds);
        return true;
    }

    private static bool TryReadMapFeatureKey(JsonElement source, out string? key) {
        if (TryReadNumber(source, out double number)) {
            key = number.ToString("G17", CultureInfo.InvariantCulture);
            return true;
        }
        return TryReadOptionalText(source, out key) && !string.IsNullOrEmpty(key);
    }

    private static MapPolygon[]? ReadMapPolygons(JsonElement source, ref int budget, bool requireAntimeridianSplit = false) {
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is < 1 or > MaximumMapRegions) return null;
        var polygons = new MapPolygon[source.GetArrayLength()];
        int index = 0;
        foreach (JsonElement item in source.EnumerateArray()) {
            MapPolygon? polygon = ReadMapPolygon(item, ref budget, requireAntimeridianSplit);
            if (polygon == null) return null;
            polygons[index++] = polygon;
        }
        return polygons;
    }

    private static MapPolygon? ReadMapPolygon(JsonElement source, ref int budget, bool requireAntimeridianSplit = false) {
        if (source.ValueKind != JsonValueKind.Array || source.GetArrayLength() is < 1 or > 128) return null;
        var rings = new MapCoordinate[source.GetArrayLength()][];
        int ringIndex = 0;
        MapBounds bounds = MapBounds.Empty;
        foreach (JsonElement ring in source.EnumerateArray()) {
            if (ring.ValueKind != JsonValueKind.Array || ring.GetArrayLength() < 4 || ring.GetArrayLength() > budget) return null;
            budget -= ring.GetArrayLength();
            var positions = new MapCoordinate[ring.GetArrayLength() - 1];
            int index = 0;
            foreach (JsonElement point in ring.EnumerateArray()) {
                if (!TryReadMapGeoJsonPoint(point, out MapCoordinate position)) return null;
                if (requireAntimeridianSplit && index > 0 && Math.Abs(position.Longitude - positions[index - 1].Longitude) > 180)
                    return null;
                if (index == positions.Length) {
                    if (positions[0] != position) return null;
                } else {
                    positions[index] = position;
                    bounds = bounds.Include(position);
                }
                index++;
            }
            rings[ringIndex++] = positions;
        }
        return new MapPolygon(rings, bounds);
    }

    private static bool TryReadMapGeoJsonPoint(JsonElement point, out MapCoordinate position) {
        position = default;
        if (point.ValueKind != JsonValueKind.Array || point.GetArrayLength() is < 2 or > 3 ||
            !TryReadNumber(point[0], out double x) || !TryReadNumber(point[1], out double y) || !IsMapPosition(x, y)) return false;
        position = new MapCoordinate(x, y);
        return true;
    }

    private static bool IsMapPosition(double longitude, double latitude) => longitude is >= -180 and <= 180 && latitude is >= -90 and <= 90;

    private static MapBounds GetMapBounds(IReadOnlyList<MapPolygon> polygons) {
        MapBounds bounds = MapBounds.Empty;
        foreach (MapPolygon polygon in polygons) bounds = bounds.Include(polygon.Bounds);
        return bounds;
    }

    private static MapWorld LoadMapWorld() {
        const string resource = "OfficeIMO.Notebook.Markdown.NaturalEarth110m.json.gz";
        using var compressed = typeof(NotebookPlotlyChart).Assembly.GetManifestResourceStream(resource)
            ?? throw new InvalidOperationException("The embedded Natural Earth map is missing.");
        using var data = new GZipStream(compressed, CompressionMode.Decompress);
        using JsonDocument json = JsonDocument.Parse(data);
        JsonElement root = json.RootElement;
        int landBudget = MaximumMapVertices;
        int countryBudget = MaximumMapVertices;
        MapPolygon[] land = ReadMapPolygons(root.GetProperty("land"), ref landBudget)
            ?? throw new InvalidOperationException("Invalid embedded Natural Earth land geometry.");
        var countries = new List<MapFeature>();
        var index = new Dictionary<string, MapFeature>(StringComparer.OrdinalIgnoreCase);
        foreach (JsonElement entry in root.GetProperty("countries").EnumerateArray()) {
            string code = entry[0].GetString()!;
            string name = entry[1].GetString()!;
            MapPolygon[] polygons = ReadMapPolygons(entry[4], ref countryBudget)
                ?? throw new InvalidOperationException("Invalid embedded Natural Earth country geometry.");
            var country = new MapFeature(code, name, new MapCoordinate(entry[2].GetDouble(), entry[3].GetDouble()),
                polygons, GetMapBounds(polygons));
            countries.Add(country);
            if (code != "-99") index.Add(code, country);
        }
        return new MapWorld(land, countries, index);
    }

    private sealed record MapPolygon(MapCoordinate[][] Rings, MapBounds Bounds);

    private sealed record MapFeature(string Key, string Name, MapCoordinate LabelPosition, MapPolygon[] Polygons, MapBounds Bounds);

    private sealed record MapRegion(MapFeature Feature, double Value, string Label);

    private sealed record MapChoropleth(List<MapRegion> Regions, MapBounds? GeoJsonBounds, SceneScale Scale,
        string? ScaleTitle, OfficeColor Border, double BorderWidth, double Opacity);

    private sealed record MapWorld(MapPolygon[] Land, IReadOnlyList<MapFeature> Countries,
        IReadOnlyDictionary<string, MapFeature> CountriesByCode);

    private readonly record struct MapBounds(double West, double South, double East, double North) {
        internal static MapBounds Empty => new(double.PositiveInfinity, double.PositiveInfinity, double.NegativeInfinity, double.NegativeInfinity);

        internal MapCoordinate Center => new((West + East) / 2, (South + North) / 2);

        internal MapBounds Include(MapCoordinate point) => new(Math.Min(West, point.Longitude), Math.Min(South, point.Latitude),
            Math.Max(East, point.Longitude), Math.Max(North, point.Latitude));

        internal MapBounds Include(MapBounds bounds) => new(Math.Min(West, bounds.West), Math.Min(South, bounds.South),
            Math.Max(East, bounds.East), Math.Max(North, bounds.North));

        internal bool Intersects(MapBounds bounds) => West <= bounds.East && East >= bounds.West && South <= bounds.North && North >= bounds.South;

        internal bool Contains(MapCoordinate point) => point.Longitude >= West && point.Longitude <= East &&
            point.Latitude >= South && point.Latitude <= North;

        internal bool Contains(MapBounds bounds) => bounds.West >= West && bounds.East <= East && bounds.South >= South && bounds.North <= North;
    }
}
