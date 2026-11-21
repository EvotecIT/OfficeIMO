using System.Globalization;
using Microsoft.Extensions.Configuration;

namespace OfficeIMO.Toml.Configuration;

/// <summary>
/// Flattens a TOML document into configuration keys: table keys and array indexes joined with <c>:</c>, and empty tables and
/// arrays as <see langword="null"/>, as the JSON provider does. Inline tables, dotted keys and arrays of tables flatten like
/// regular tables.
/// </summary>
internal sealed class TomlConfigurationFileParser {
    private readonly Dictionary<string, string?> _data = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    private TomlConfigurationFileParser() { }

    public static IDictionary<string, string?> Parse(Stream input) {
        TomlDocument document;
        try {
            document = TomlDocument.Load(input);
        } catch (TomlException exception) {
            throw new FormatException(string.Format(CultureInfo.InvariantCulture,
                "Could not parse the TOML configuration at line {0}, column {1}: {2}", exception.Line, exception.Column, exception.Message), exception);
        }

        using (document) {
            var parser = new TomlConfigurationFileParser();
            parser.VisitTable(document.Root, path: null);
            return parser._data;
        }
    }

    private void VisitNode(TomlNode node, string path) {
        switch (node) {
            case TomlTable table:
                VisitTable(table, path);
                break;
            case TomlTableArray tables:
                if (tables.Count == 0) {
                    Set(path, null, tables);
                    break;
                }
                for (int i = 0; i < tables.Count; i++) VisitTable(tables[i], Combine(path, i));
                break;
            case TomlArray array:
                if (array.Count == 0) {
                    Set(path, null, array);
                    break;
                }
                for (int i = 0; i < array.Count; i++) VisitNode(array[i], Combine(path, i));
                break;
            case TomlValue value:
                Set(path, GetText(value), value);
                break;
        }
    }

    // The root table has no path; an empty quoted key ("") is a real segment, so a null path marks the root.
    private void VisitTable(TomlTable table, string? path) {
        // TOML keys are case-sensitive and the document already rejects exact duplicates; configuration keys are not, so
        // keys that differ only in case would silently overwrite each other.
        var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (KeyValuePair<string, TomlNode> entry in table) {
            string key = Combine(path, entry.Key);
            if (!keys.Add(entry.Key)) throw Duplicate(key, entry.Value);
            VisitNode(entry.Value, key);
        }
        if (keys.Count == 0 && path is not null) Set(path, null, table);
    }

    // Different TOML keys can flatten to one configuration key ("a:b" = 1 beside a = { b = [] }), so every write checks.
    private void Set(string path, string? value, TomlNode node) {
        if (_data.ContainsKey(path)) throw Duplicate(path, node);
        _data[path] = value;
    }

    // Strings keep their decoded text. Numbers keep their spelling where .NET parses it (decimal digits, exponents) and are
    // otherwise written as invariant decimals, so 0xFF, 1_000, inf and nan bind to numeric types. Date-times become ISO 8601.
    private static string GetText(TomlValue value) {
        switch (value.Kind) {
            case TomlNodeKind.String:
                return (string)value.Value;
            case TomlNodeKind.Integer: {
                string text = value.AsString();
                return IsDecimalInteger(text) ? text : ((long)value.Value).ToString(CultureInfo.InvariantCulture);
            }
            case TomlNodeKind.Float: {
                double number = (double)value.Value;
                if (double.IsNaN(number)) return "NaN";
                if (double.IsPositiveInfinity(number)) return "Infinity";
                if (double.IsNegativeInfinity(number)) return "-Infinity";
                string text = value.AsString();
                return text.IndexOf('_') < 0 && text[0] != '+' ? text : OfficeIMO.OfficeRoundTripNumberText.Format(number, CultureInfo.InvariantCulture);
            }
            case TomlNodeKind.Boolean:
                return (bool)value.Value ? "true" : "false";
            default:
                return ((TomlDateTime)value.Value).ToString();
        }
    }

    private static bool IsDecimalInteger(string text) {
        int start = text.Length > 1 && text[0] == '-' ? 1 : 0;
        for (int i = start; i < text.Length; i++) {
            if (text[i] < '0' || text[i] > '9') return false;
        }
        return text.Length > start;
    }

    private static string Combine(string? path, int index) => Combine(path, index.ToString(CultureInfo.InvariantCulture));

    private static string Combine(string? path, string key) => path is null ? key : path + ConfigurationPath.KeyDelimiter + key;

    private static FormatException Duplicate(string key, TomlNode node) {
        int line = node.KeyStart.Line > 0 ? node.KeyStart.Line : node.Start.Line;
        return new FormatException(string.Format(CultureInfo.InvariantCulture, "A duplicate key '{0}' was found at line {1}.", key, line));
    }
}
