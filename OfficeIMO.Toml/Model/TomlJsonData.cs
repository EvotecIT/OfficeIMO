using OfficeIMO.Json;

namespace OfficeIMO.Toml;

/// <summary>
/// Converts the document model to the JSON data model that <see cref="JsonSchema"/> validates: tables become objects,
/// arrays and arrays of tables become arrays, dates and times become their RFC 3339 text, and every value keeps its
/// source position and node.
/// </summary>
internal static class TomlJsonData {
    private const int MaxDepth = 512;

    public static JsonData Convert(TomlNode node) => Convert(node, 0);

    // Edited documents can nest deeper than the reader allows, so the depth is checked here too.
    private static JsonData Convert(TomlNode node, int depth) {
        if (depth > MaxDepth) throw new InvalidOperationException("The TOML value nests deeper than " + MaxDepth.ToString(System.Globalization.CultureInfo.InvariantCulture) + " levels.");
        TomlPosition start = node.Start;
        var location = new JsonDataLocation(start.Line, start.Column, node);
        switch (node) {
            case TomlTable table: {
                var properties = new List<KeyValuePair<string, JsonData>>(table.Count);
                foreach (KeyValuePair<string, TomlNode> entry in table) properties.Add(new KeyValuePair<string, JsonData>(entry.Key, Convert(entry.Value, depth + 1)));
                return JsonData.CreateObject(properties, location);
            }
            case TomlTableArray tables: {
                var items = new List<JsonData>(tables.Count);
                foreach (TomlTable item in tables) items.Add(Convert(item, depth + 1));
                return JsonData.CreateArray(items, location);
            }
            case TomlArray array: {
                var items = new List<JsonData>(array.Count);
                foreach (TomlNode item in array) items.Add(Convert(item, depth + 1));
                return JsonData.CreateArray(items, location);
            }
            default: {
                object value = ((TomlValue)node).Value;
                return value switch {
                    string text => JsonData.CreateString(text, location),
                    long integer => JsonData.CreateNumber(integer, location),
                    double number => JsonData.CreateNumber(number, location),
                    bool flag => JsonData.CreateBoolean(flag, location),
                    _ => JsonData.CreateString(value.ToString() ?? string.Empty, location)
                };
            }
        }
    }
}
