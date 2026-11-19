using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Toml.Benchmarks;

/// <summary>
/// Content-level views of each library's result, used to validate the measured lanes before timing: the ordered keys
/// and decoded scalar values an event parser reports, and the leaf values, with their paths, a document model holds.
/// Scalars are written as <c>kind:value</c> (<c>s</c>tring, <c>i</c>nteger, <c>f</c>loat, <c>b</c>oolean, <c>d</c>ate-time).
/// </summary>
internal static class TomlContentCheck {
    // ---- Event sequences ------------------------------------------------------------------------------------------

    /// <summary>Keys (<c>k:</c>) and scalar values in the order OfficeIMO's pull reader reports them.</summary>
    public static List<string> EventsOfficeIMO(string toml) {
        var events = new List<string>();
        var reader = new TomlReader(toml);
        while (reader.Read()) {
            switch (reader.TokenType) {
                case TomlTokenType.Key: events.Add("k:" + reader.GetString()); break;
                case TomlTokenType.String: events.Add(Scalar(reader.GetString())); break;
                case TomlTokenType.Integer: events.Add(Scalar(reader.GetInt64())); break;
                case TomlTokenType.Float: events.Add(Scalar(reader.GetDouble())); break;
                case TomlTokenType.Boolean: events.Add(Scalar(reader.GetBoolean())); break;
                case TomlTokenType.OffsetDateTime:
                case TomlTokenType.LocalDateTime:
                case TomlTokenType.LocalDate:
                case TomlTokenType.LocalTime:
                    events.Add("d:" + reader.GetDateTime().ToString());
                    break;
            }
        }
        return events;
    }

    /// <summary>Property names (<c>k:</c>) and scalar values in the order Tomlyn's event parser reports them.</summary>
    public static List<string> EventsTomlyn(string toml) {
        var events = new List<string>();
        var parser = Tomlyn.Parsing.TomlParser.Create(toml);
        while (parser.MoveNext()) {
            switch (parser.Current.Kind) {
                case Tomlyn.Parsing.TomlParseEventKind.PropertyName: events.Add("k:" + parser.GetPropertyName()); break;
                case Tomlyn.Parsing.TomlParseEventKind.String: events.Add(Scalar(parser.GetString())); break;
                case Tomlyn.Parsing.TomlParseEventKind.Integer: events.Add(Scalar(parser.Current.GetInt64())); break;
                case Tomlyn.Parsing.TomlParseEventKind.Float: events.Add(Scalar(parser.Current.GetDouble())); break;
                case Tomlyn.Parsing.TomlParseEventKind.Boolean: events.Add(Scalar(parser.Current.GetBoolean())); break;
                case Tomlyn.Parsing.TomlParseEventKind.DateTime: events.Add("d:" + parser.Current.GetTomlDateTime().ToString()); break;
            }
        }
        return events;
    }

    // ---- Document leaves ------------------------------------------------------------------------------------------

    /// <summary>Every leaf of an OfficeIMO document as <c>path=kind:value</c>, sorted.</summary>
    public static List<string> LeavesOfficeIMO(TomlDocument document) {
        var leaves = new List<string>();
        AddOfficeIMO(document.Root, string.Empty, leaves);
        return Sorted(leaves);
    }

    public static List<string> LeavesTomlyn(Tomlyn.Model.TomlTable table) {
        var leaves = new List<string>();
        AddTomlyn(table, string.Empty, leaves);
        return Sorted(leaves);
    }

    public static List<string> LeavesTomlet(Tomlet.Models.TomlDocument document) {
        var leaves = new List<string>();
        AddTomlet(document, string.Empty, leaves);
        return Sorted(leaves);
    }

    /// <summary>Every leaf of a CsToml document, read through its own dictionary projection.</summary>
    public static List<string> LeavesCsToml(CsToml.TomlDocument document) {
        var leaves = new List<string>();
        AddCsToml(document.ToDictionary<object, object>(), string.Empty, leaves);
        return Sorted(leaves);
    }

    /// <summary>The scalar part (<c>kind:value</c>) of each leaf, sorted: the values a parser must report.</summary>
    public static List<string> ScalarsOf(IEnumerable<string> leaves) =>
        Sorted(leaves.Select(leaf => leaf.Substring(leaf.IndexOf('\u001E') + 1)));

    /// <summary>
    /// Each value of an event sequence with the key that names it (the last key read before it), in document order.
    /// Table headers are left out: the parsers spell a reopened parent table differently, while the key of every value
    /// and the values themselves must agree.
    /// </summary>
    public static List<string> KeyedValues(IEnumerable<string> events) {
        var values = new List<string>();
        string key = string.Empty;
        foreach (string item in events) {
            if (item.StartsWith("k:", StringComparison.Ordinal)) {
                key = item;
            } else {
                values.Add(key + '' + item);
            }
        }
        return values;
    }

    /// <summary>The scalar entries of an event sequence, sorted.</summary>
    public static List<string> ScalarsOfEvents(IEnumerable<string> events) =>
        Sorted(events.Where(item => !item.StartsWith("k:", StringComparison.Ordinal)));

    public static string Hash(IEnumerable<string> items) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(string.Join("\n", items))))[..16];

    /// <summary>Runs one lane's check, then records it under the library's declared id and case.</summary>
    public static void Record(string library, string caseName, string detail, Func<string?> check) {
        string? failure;
        try {
            failure = check();
        } catch (Exception exception) when (exception is InvalidOperationException or FormatException or ArgumentException) {
            failure = exception.GetType().Name + ": " + exception.Message;
        }
        if (failure is not null) {
            BenchmarkOutputValidationLog.Failed(library, detail + "; " + failure, caseName);
            throw new InvalidOperationException(library + " (" + caseName + "): " + failure);
        }
        BenchmarkOutputValidationLog.Passed(library, detail, caseName);
    }

    /// <summary>Returns null when both lists match, otherwise where they first differ.</summary>
    public static string? Compare(string what, IReadOnlyList<string> expected, IReadOnlyList<string> actual) {
        int shared = Math.Min(expected.Count, actual.Count);
        for (int index = 0; index < shared; index++) {
            if (!string.Equals(expected[index], actual[index], StringComparison.Ordinal)) {
                return what + " differ at item " + index + ": expected '" + Clip(expected[index]) + "' but got '" + Clip(actual[index]) + "'";
            }
        }
        return expected.Count == actual.Count ? null : what + ": expected " + expected.Count + " items but got " + actual.Count;
    }

    private static string Clip(string value) {
        string readable = value.Replace('\u001F', '.').Replace('\u001E', '=');
        return readable.Length <= 160 ? readable : readable.Substring(0, 160) + "...";
    }

    private static void AddOfficeIMO(TomlNode node, string path, List<string> leaves) {
        switch (node) {
            case TomlTable table:
                foreach (KeyValuePair<string, TomlNode> entry in table) AddOfficeIMO(entry.Value, Child(path, entry.Key), leaves);
                return;
            case TomlTableArray tables: {
                int index = 0;
                foreach (TomlTable table in tables) AddOfficeIMO(table, Item(path, index++), leaves);
                return;
            }
            case TomlArray array: {
                int index = 0;
                foreach (TomlNode item in array) AddOfficeIMO(item, Item(path, index++), leaves);
                return;
            }
            default:
                leaves.Add(Leaf(path, ((TomlValue)node).Value));
                return;
        }
    }

    private static void AddTomlyn(object? value, string path, List<string> leaves) {
        switch (value) {
            case Tomlyn.Model.TomlTable table:
                foreach (KeyValuePair<string, object> entry in table) AddTomlyn(entry.Value, Child(path, entry.Key), leaves);
                return;
            case Tomlyn.Model.TomlTableArray tables: {
                int index = 0;
                foreach (Tomlyn.Model.TomlTable table in tables) AddTomlyn(table, Item(path, index++), leaves);
                return;
            }
            case Tomlyn.Model.TomlArray array: {
                int index = 0;
                foreach (object? item in array) AddTomlyn(item, Item(path, index++), leaves);
                return;
            }
            default:
                leaves.Add(Leaf(path, value));
                return;
        }
    }

    private static void AddTomlet(Tomlet.Models.TomlValue value, string path, List<string> leaves) {
        switch (value) {
            case Tomlet.Models.TomlTable table:
                foreach (KeyValuePair<string, Tomlet.Models.TomlValue> entry in table.Entries) AddTomlet(entry.Value, Child(path, entry.Key), leaves);
                return;
            case Tomlet.Models.TomlArray array: {
                int index = 0;
                foreach (Tomlet.Models.TomlValue item in array.ArrayValues) AddTomlet(item, Item(path, index++), leaves);
                return;
            }
            case Tomlet.Models.TomlString text: leaves.Add(Leaf(path, text.Value)); return;
            case Tomlet.Models.TomlLong integer: leaves.Add(Leaf(path, integer.Value)); return;
            case Tomlet.Models.TomlDouble number: leaves.Add(Leaf(path, number.Value)); return;
            case Tomlet.Models.TomlBoolean boolean: leaves.Add(Leaf(path, boolean.Value)); return;
            default: leaves.Add(path + '\u001E' + "d:" + value.StringValue); return;
        }
    }

    private static void AddCsToml(object? value, string path, List<string> leaves) {
        switch (value) {
            case System.Collections.IDictionary table:
                foreach (System.Collections.DictionaryEntry entry in table) AddCsToml(entry.Value, Child(path, Convert.ToString(entry.Key, CultureInfo.InvariantCulture)!), leaves);
                return;
            case string:
                leaves.Add(Leaf(path, value));
                return;
            case System.Collections.IEnumerable array: {
                int index = 0;
                foreach (object? item in array) AddCsToml(item, Item(path, index++), leaves);
                return;
            }
            default:
                leaves.Add(Leaf(path, value));
                return;
        }
    }

    private static string Child(string path, string key) => path.Length == 0 ? key : path + '\u001F' + key;

    private static string Item(string path, int index) => path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";

    private static string Leaf(string path, object? value) => path + '\u001E' + Scalar(value);

    private static string Scalar(object? value) => value switch {
        string text => "s:" + text,
        long or int => "i:" + Convert.ToString(value, CultureInfo.InvariantCulture),
        double number => "f:" + (double.IsNaN(number) ? "nan" : number.ToString("R", CultureInfo.InvariantCulture)),
        bool boolean => boolean ? "b:true" : "b:false",
        null => "null",
        _ => "d:" + Convert.ToString(value, CultureInfo.InvariantCulture)
    };

    private static List<string> Sorted(IEnumerable<string> items) {
        var list = items.ToList();
        list.Sort(StringComparer.Ordinal);
        return list;
    }
}
