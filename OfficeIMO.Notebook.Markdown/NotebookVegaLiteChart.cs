using System.Globalization;
using System.Text.Json;
using System.Xml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Notebook.Markdown;

internal static class NotebookVegaLiteChart {
    private const int MaximumPoints = 256;
    private const int MaximumCategories = 48;
    private const int MaximumTextCharacters = 128;

    internal static OfficeChartSnapshot? TryRead(JsonElement spec, CancellationToken cancellationToken) {
        if (!HasOnly(spec, "$schema", "data", "datasets", "mark", "encoding", "title", "width", "height", "config", "description")
            || !spec.TryGetProperty("mark", out JsonElement mark)
            || !TryReadMark(mark, out string kind, out OfficeColor? color)
            || !TryFindRows(spec, out JsonElement rows)
            || !spec.TryGetProperty("encoding", out JsonElement encoding)
            || !HasOnly(encoding, "x", "y", "tooltip")
            || !TryReadAxis(encoding, "x", out Axis x)
            || !TryReadAxis(encoding, "y", out Axis y)
            || !TryTitle(spec, "title", out string? title)
            || !TrySize(spec, out double width, out double height)) return null;

        bool horizontal = x.IsNumeric && !y.IsNumeric;
        bool vertical = !x.IsNumeric && y.IsNumeric;
        if (kind == "bar") {
            if (!horizontal && !vertical || rows.GetArrayLength() > MaximumCategories) return null;
            return ReadBars(rows, horizontal ? y : x, horizontal ? x : y, horizontal, title, color, width, height, cancellationToken);
        }
        if (!x.IsNumeric || !y.IsNumeric || kind == "line" && rows.GetArrayLength() < 2) return null;
        return ReadPoints(rows, x, y, kind == "line", title, color, width, height, cancellationToken);
    }

    private static OfficeChartSnapshot? ReadBars(JsonElement rows, Axis categoryAxis, Axis valueAxis, bool horizontal,
        string? title, OfficeColor? color, double width, double height, CancellationToken cancellationToken) {
        var keys = new HashSet<string>(StringComparer.Ordinal);
        var data = new List<(string Category, double Value)>(rows.GetArrayLength());
        foreach (JsonElement row in rows.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryGetValue(row, categoryAxis.Field, out JsonElement categoryValue)
                || !TryText(categoryValue, out string category) || !keys.Add(category)
                || !TryNumber(row, valueAxis.Field, out double value)) return null;
            data.Add((category, value));
        }
        if (categoryAxis.Order != null) {
            if (categoryAxis.Order.Count != keys.Count || !keys.All(categoryAxis.Order.ContainsKey)) return null;
            data.Sort((left, right) => categoryAxis.Order[left.Category].CompareTo(categoryAxis.Order[right.Category]));
        } else if (categoryAxis.Sort) {
            data.Sort((left, right) => StringComparer.Ordinal.Compare(left.Category, right.Category));
        }
        var series = new OfficeSeries(valueAxis.Title ?? valueAxis.Field, data.Select(point => point.Value), null, color);
        var style = new OfficeChartStyle(showBorder: false);
        var layout = new OfficeChartLayout(showLegend: false, categoryAxisTitle: categoryAxis.Title,
            valueAxisTitle: valueAxis.Title, reverseCategoryAxis: horizontal);
        return new OfficeChartSnapshot("Vega-Lite", title, horizontal ? OfficeChartType.Bar : OfficeChartType.Column,
            new OfficeChartData(data.Select(point => point.Category), new[] { series }),
            OfficeLength.Pixels(width), OfficeLength.Pixels(height), style, layout);
    }

    private static OfficeChartSnapshot? ReadPoints(JsonElement rows, Axis x, Axis y, bool connect,
        string? title, OfficeColor? color, double width, double height, CancellationToken cancellationToken) {
        var points = new List<(double X, double Y)>(rows.GetArrayLength());
        foreach (JsonElement row in rows.EnumerateArray()) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryNumber(row, x.Field, out double xValue) || !TryNumber(row, y.Field, out double yValue)) return null;
            points.Add((xValue, yValue));
        }
        if (connect) points = points.OrderBy(point => point.X).ToList();
        var series = new OfficeSeries(y.Title ?? y.Field, points.Select(point => point.Y), points.Select(point => point.X),
            color, pointColors: null, showMarkers: !connect, connectLine: connect);
        var style = new OfficeChartStyle(showBorder: false);
        var layout = new OfficeChartLayout(showLegend: false, categoryAxisTitle: x.Title, valueAxisTitle: y.Title,
            showMarkers: !connect, connectScatterPoints: connect);
        return new OfficeChartSnapshot("Vega-Lite", title, OfficeChartType.Scatter,
            new OfficeChartData(points.Select(point => point.X.ToString("R", CultureInfo.InvariantCulture)), new[] { series }),
            OfficeLength.Pixels(width), OfficeLength.Pixels(height), style, layout);
    }

    private static bool TryReadMark(JsonElement mark, out string kind, out OfficeColor? color) {
        color = null;
        if (mark.ValueKind == JsonValueKind.Object) {
            if (!HasOnly(mark, "type", "color", "tooltip", "filled")
                || !mark.TryGetProperty("type", out JsonElement type)
                || !TryText(type, out kind)) {
                kind = string.Empty;
                return false;
            }
            if (mark.TryGetProperty("color", out JsonElement paint)) {
                if (!TryText(paint, out string paintText) || !OfficeColor.TryParseCss(paintText, out OfficeColor parsed)) return false;
                color = parsed;
            }
        } else if (!TryText(mark, out kind)) {
            return false;
        }
        return kind is "bar" or "line" or "point" or "circle";
    }

    private static bool TryFindRows(JsonElement spec, out JsonElement rows) {
        rows = default;
        if (!spec.TryGetProperty("data", out JsonElement data) || !HasOnly(data, "values", "name")) return false;
        if (data.TryGetProperty("name", out JsonElement name)) {
            if (data.TryGetProperty("values", out _) || !TryText(name, out string dataset)
                || !spec.TryGetProperty("datasets", out JsonElement datasets)
                || !TryGetValue(datasets, dataset, out rows)) return false;
        } else if (!data.TryGetProperty("values", out rows)) {
            return false;
        }
        return rows.ValueKind == JsonValueKind.Array && rows.GetArrayLength() is > 0 and <= MaximumPoints;
    }

    private static bool TryReadAxis(JsonElement encoding, string name, out Axis axis) {
        axis = default;
        if (!encoding.TryGetProperty(name, out JsonElement value)
            || !HasOnly(value, "field", "type", "title", "sort")
            || !value.TryGetProperty("field", out JsonElement fieldValue) || !TryText(fieldValue, out string field)
            || !value.TryGetProperty("type", out JsonElement typeValue) || !TryText(typeValue, out string type)
            || type is not ("quantitative" or "nominal" or "ordinal")
            || !TryTitle(value, "title", out string? title)) return false;
        // Vega field paths have a different meaning from plain JSON property names.
        if (field.IndexOfAny(new[] { '.', '[', ']', '\\' }) >= 0) return false;
        bool hasSort = value.TryGetProperty("sort", out JsonElement sort);
        if (!TryReadOrder(sort, out Dictionary<string, int>? order) || hasSort && type == "quantitative") return false;
        axis = new Axis(field, value.TryGetProperty("title", out _) ? title : field,
            type == "quantitative", sort.ValueKind != JsonValueKind.Null, order);
        return true;
    }

    private static bool TryReadOrder(JsonElement sort, out Dictionary<string, int>? order) {
        order = null;
        if (sort.ValueKind is JsonValueKind.Undefined or JsonValueKind.Null) return true;
        if (sort.ValueKind == JsonValueKind.String) return sort.ValueEquals("ascending");
        if (sort.ValueKind != JsonValueKind.Array || sort.GetArrayLength() is 0 or > MaximumCategories) return false;
        order = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (JsonElement item in sort.EnumerateArray()) {
            if (!TryText(item, out string category) || order.ContainsKey(category)) return false;
            order.Add(category, order.Count);
        }
        return true;
    }

    private static bool TrySize(JsonElement spec, out double width, out double height) {
        width = 520;
        height = 320;
        if (spec.TryGetProperty("config", out JsonElement config)) {
            if (!HasOnly(config, "view") || !config.TryGetProperty("view", out JsonElement view)
                || !HasOnly(view, "continuousWidth", "continuousHeight")
                || !ReadSize(view, "continuousWidth", ref width)
                || !ReadSize(view, "continuousHeight", ref height)) return false;
        }
        return ReadSize(spec, "width", ref width) && ReadSize(spec, "height", ref height);
    }

    private static bool ReadSize(JsonElement spec, string property, ref double size) {
        if (!spec.TryGetProperty(property, out JsonElement value)) return true;
        return value.ValueKind == JsonValueKind.Number && value.TryGetDouble(out size) && size is >= 120 and <= 1200;
    }

    private static bool TryTitle(JsonElement item, string property, out string? title) {
        title = null;
        if (!item.TryGetProperty(property, out JsonElement value) || value.ValueKind == JsonValueKind.Null) return true;
        if (!TryText(value, out string text)) return false;
        title = text;
        return true;
    }

    private static bool TryText(JsonElement element, out string text) {
        text = string.Empty;
        if (element.ValueKind != JsonValueKind.String) return false;
        try {
            string value = element.GetString() ?? string.Empty;
            if (value.Length is 0 or > MaximumTextCharacters || value.Any(char.IsControl)) return false;
            text = XmlConvert.VerifyXmlChars(value);
            return true;
        } catch (InvalidOperationException) {
            return false;
        } catch (XmlException) {
            return false;
        }
    }

    private static bool TryNumber(JsonElement row, string property, out double number) {
        number = 0;
        return TryGetValue(row, property, out JsonElement value) && value.ValueKind == JsonValueKind.Number
            && value.TryGetDouble(out number) && number is >= -1_000_000_000 and <= 1_000_000_000
            && (number == 0 || Math.Abs(number) >= 0.000_000_001);
    }

    private static bool TryGetValue(JsonElement item, string property, out JsonElement value) {
        value = default;
        return item.ValueKind == JsonValueKind.Object && item.TryGetProperty(property, out value);
    }

    private static bool HasOnly(JsonElement item, params string[] names) =>
        item.ValueKind == JsonValueKind.Object
        && item.EnumerateObject().All(property => names.Any(name => property.NameEquals(name)));

    private readonly struct Axis {
        internal Axis(string field, string? title, bool isNumeric, bool sort, IReadOnlyDictionary<string, int>? order) {
            Field = field;
            Title = title;
            IsNumeric = isNumeric;
            Sort = sort;
            Order = order;
        }

        internal string Field { get; }
        internal string? Title { get; }
        internal bool IsNumeric { get; }
        internal bool Sort { get; }
        internal IReadOnlyDictionary<string, int>? Order { get; }
    }
}
