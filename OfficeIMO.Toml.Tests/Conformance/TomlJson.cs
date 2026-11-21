using System.Globalization;
using System.Text;

namespace OfficeIMO.Toml.Tests.Conformance;

/// <summary>A parsed JSON value: an object, array, string or other literal. Enough for toml-test files and manifests.</summary>
public sealed class JsonNode {
    public Dictionary<string, JsonNode>? Object { get; set; }
    public List<JsonNode>? Array { get; set; }
    public string? String { get; set; }
    public string? Literal { get; set; }

    public JsonNode? Get(string key) => Object != null && Object.TryGetValue(key, out JsonNode? value) ? value : null;

    public string GetString(string key) => Get(key)?.String ?? throw new InvalidDataException("Missing string '" + key + "'.");
}

/// <summary>A strict, minimal JSON reader, so the harness has no dependency that net472 would need to restore.</summary>
public static class JsonText {
    public static JsonNode Parse(string text) {
        int i = 0;
        JsonNode node = Value(text, ref i);
        Skip(text, ref i);
        if (i != text.Length) throw new InvalidDataException("Trailing JSON at " + i + ".");
        return node;
    }

    private static JsonNode Value(string s, ref int i) {
        Skip(s, ref i);
        char ch = s[i];
        if (ch == '{') {
            var map = new Dictionary<string, JsonNode>(StringComparer.Ordinal);
            i++;
            Skip(s, ref i);
            if (s[i] == '}') {
                i++;
                return new JsonNode { Object = map };
            }
            while (true) {
                Skip(s, ref i);
                string key = StringValue(s, ref i);
                Skip(s, ref i);
                Expect(s, ref i, ':');
                map[key] = Value(s, ref i);
                Skip(s, ref i);
                if (s[i] == ',') {
                    i++;
                    continue;
                }
                Expect(s, ref i, '}');
                return new JsonNode { Object = map };
            }
        }
        if (ch == '[') {
            var list = new List<JsonNode>();
            i++;
            Skip(s, ref i);
            if (s[i] == ']') {
                i++;
                return new JsonNode { Array = list };
            }
            while (true) {
                list.Add(Value(s, ref i));
                Skip(s, ref i);
                if (s[i] == ',') {
                    i++;
                    continue;
                }
                Expect(s, ref i, ']');
                return new JsonNode { Array = list };
            }
        }
        if (ch == '"') return new JsonNode { String = StringValue(s, ref i) };
        int start = i;
        while (i < s.Length && ",}] \t\r\n".IndexOf(s[i]) < 0) i++;
        return new JsonNode { Literal = s.Substring(start, i - start) };
    }

    private static string StringValue(string s, ref int i) {
        Expect(s, ref i, '"');
        var text = new StringBuilder();
        while (true) {
            char ch = s[i++];
            if (ch == '"') return text.ToString();
            if (ch != '\\') {
                text.Append(ch);
                continue;
            }
            char escape = s[i++];
            switch (escape) {
                case 'b': text.Append('\b'); break;
                case 'f': text.Append('\f'); break;
                case 'n': text.Append('\n'); break;
                case 'r': text.Append('\r'); break;
                case 't': text.Append('\t'); break;
                case 'u':
                    text.Append((char)int.Parse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    i += 4;
                    break;
                default: text.Append(escape); break;
            }
        }
    }

    private static void Skip(string s, ref int i) {
        while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
    }

    private static void Expect(string s, ref int i, char ch) {
        if (s[i] != ch) throw new InvalidDataException("Expected '" + ch + "' at " + i + ".");
        i++;
    }
}

/// <summary>Compares a loaded document with a toml-test expectation in its tagged JSON form.</summary>
public static class TomlTestJson {
    /// <summary>Returns null when equal, otherwise the path and reason of the first difference.</summary>
    public static string? Compare(TomlNode actual, JsonNode expected, string path = "$") {
        switch (actual) {
            case TomlTable table:
                if (expected.Object is null || IsLeaf(expected)) return path + ": expected " + Describe(expected) + " but found a table";
                if (expected.Object.Count != table.Count) return path + ": expected " + expected.Object.Count + " keys but found " + table.Count + " (" + string.Join(", ", table.Keys) + ")";
                foreach (KeyValuePair<string, TomlNode> entry in table) {
                    if (!expected.Object.TryGetValue(entry.Key, out JsonNode? value)) return path + ": unexpected key '" + entry.Key + "'";
                    string? difference = Compare(entry.Value, value, path + "." + entry.Key);
                    if (difference != null) return difference;
                }
                return null;
            case TomlTableArray tables:
                return CompareItems(tables.Cast<TomlNode>().ToList(), expected, path);
            case TomlArray array:
                return CompareItems(array.ToList(), expected, path);
            case TomlValue value:
                if (!IsLeaf(expected)) return path + ": expected " + Describe(expected) + " but found a " + value.Kind;
                return CompareValue(value, expected.GetString("type"), expected.GetString("value"), path);
            default:
                return path + ": unknown node";
        }
    }

    private static string? CompareItems(List<TomlNode> items, JsonNode expected, string path) {
        if (expected.Array is null) return path + ": expected " + Describe(expected) + " but found an array";
        if (expected.Array.Count != items.Count) return path + ": expected " + expected.Array.Count + " items but found " + items.Count;
        for (int i = 0; i < items.Count; i++) {
            string? difference = Compare(items[i], expected.Array[i], path + "[" + i + "]");
            if (difference != null) return difference;
        }
        return null;
    }

    private static bool IsLeaf(JsonNode node) =>
        node.Object != null && node.Object.Count == 2 && node.Get("type")?.String != null && node.Get("value")?.String != null;

    private static string Describe(JsonNode node) =>
        IsLeaf(node) ? "a " + node.GetString("type") : node.Object != null ? "a table" : node.Array != null ? "an array" : "a literal";

    private static string? CompareValue(TomlValue actual, string type, string expected, string path) {
        string? actualType = actual.Kind switch {
            TomlNodeKind.String => "string",
            TomlNodeKind.Integer => "integer",
            TomlNodeKind.Float => "float",
            TomlNodeKind.Boolean => "bool",
            _ => ((TomlDateTime)actual.Value).Kind switch {
                TomlDateTimeKind.OffsetDateTime => "datetime",
                TomlDateTimeKind.LocalDateTime => "datetime-local",
                TomlDateTimeKind.LocalDate => "date-local",
                _ => "time-local"
            }
        };
        if (actualType != type) return path + ": expected " + type + " but found " + actualType;
        bool equal = type switch {
            "string" => (string)actual.Value == expected,
            "integer" => (long)actual.Value == long.Parse(expected, CultureInfo.InvariantCulture),
            "float" => FloatEquals((double)actual.Value, expected),
            "bool" => ((bool)actual.Value ? "true" : "false") == expected,
            _ => Canonical((TomlDateTime)actual.Value) == CanonicalExpected(expected)
        };
        return equal ? null : path + ": expected " + type + " '" + expected + "' but found '" + actual.AsString() + "'";
    }

    private static bool FloatEquals(double actual, string expected) {
        string text = expected.TrimStart('+');
        if (text is "nan" or "-nan") return double.IsNaN(actual);
        if (text == "inf") return double.IsPositiveInfinity(actual);
        if (text == "-inf") return double.IsNegativeInfinity(actual);
        return actual.Equals(double.Parse(text, NumberStyles.Float, CultureInfo.InvariantCulture));
    }

    /// <summary>Formats a date-time independently of the engine's formatter: the instant in UTC for offsets, nanoseconds without trailing zeros.</summary>
    private static string Canonical(TomlDateTime value) {
        if (value.Kind == TomlDateTimeKind.OffsetDateTime) {
            DateTimeOffset utc = new DateTimeOffset(value.Year, value.Month, value.Day, value.Hour, value.Minute, Math.Min(value.Second, 59), TimeSpan.Zero)
                .AddSeconds(value.Second == 60 ? 1 : 0)
                .AddMinutes(-value.Offset!.Value.TotalMinutes);
            return utc.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture) + Fraction(value.Nanosecond) + "Z";
        }
        string date = value.Year.ToString("D4", CultureInfo.InvariantCulture) + "-" + value.Month.ToString("D2", CultureInfo.InvariantCulture) + "-" + value.Day.ToString("D2", CultureInfo.InvariantCulture);
        string time = value.Hour.ToString("D2", CultureInfo.InvariantCulture) + ":" + value.Minute.ToString("D2", CultureInfo.InvariantCulture) + ":" + value.Second.ToString("D2", CultureInfo.InvariantCulture) + Fraction(value.Nanosecond);
        return value.Kind switch {
            TomlDateTimeKind.LocalDate => date,
            TomlDateTimeKind.LocalTime => time,
            _ => date + "T" + time
        };
    }

    private static string CanonicalExpected(string text) {
        // toml-test writes RFC 3339; reduce it to the same canonical form with a small independent parser.
        text = text.Replace(' ', 'T').Replace('t', 'T').Replace('z', 'Z');
        if (text.Length >= 3 && text[2] == ':') return Time(text, 0, out _);
        string date = text.Substring(0, 10);
        if (text.Length == 10) return date;
        string time = Time(text, 11, out int end);
        if (end == text.Length) return date + "T" + time;
        int offset = 0;
        if (text[end] != 'Z') {
            int sign = text[end] == '-' ? -1 : 1;
            offset = sign * (int.Parse(text.Substring(end + 1, 2), CultureInfo.InvariantCulture) * 60 + int.Parse(text.Substring(end + 4, 2), CultureInfo.InvariantCulture));
        }
        string[] parts = time.Split('.');
        DateTimeOffset local = DateTimeOffset.ParseExact(date + "T" + parts[0], "yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal).AddMinutes(-offset);
        return local.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ss", CultureInfo.InvariantCulture) + (parts.Length > 1 ? "." + parts[1] : string.Empty) + "Z";
    }

    private static string Time(string text, int start, out int end) {
        int i = start + 5;
        string time = text.Substring(start, 5);
        string seconds = ":00";
        string fraction = string.Empty;
        if (i < text.Length && text[i] == ':') {
            seconds = text.Substring(i, 3);
            i += 3;
            if (i < text.Length && text[i] == '.') {
                int f = i + 1;
                while (f < text.Length && char.IsDigit(text[f])) f++;
                string digits = text.Substring(i + 1, f - i - 1);
                if (digits.Length > 9) digits = digits.Substring(0, 9);
                digits = digits.TrimEnd('0');
                fraction = digits.Length > 0 ? "." + digits : string.Empty;
                i = f;
            }
        }
        end = i;
        return time + seconds + fraction;
    }

    private static string Fraction(int nanoseconds) {
        if (nanoseconds == 0) return string.Empty;
        return "." + nanoseconds.ToString("D9", CultureInfo.InvariantCulture).TrimEnd('0');
    }
}
