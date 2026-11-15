using System.Collections.ObjectModel;
using System.Globalization;
using System.Text;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Json;

/// <summary>The kind of a <see cref="JsonData"/> value.</summary>
public enum JsonDataKind : byte {
    /// <summary>The JSON <c>null</c> value.</summary>
    Null = 0,
    /// <summary><c>true</c> or <c>false</c>.</summary>
    Boolean,
    /// <summary>A number, integer or not.</summary>
    Number,
    /// <summary>A string.</summary>
    String,
    /// <summary>An ordered list of values.</summary>
    Array,
    /// <summary>A set of named values in source order.</summary>
    Object
}

/// <summary>
/// An immutable value in the JSON data model: what a JSON, YAML or TOML document means, independent of its syntax. Each
/// value can carry the one-based line and column it came from and the source node, so findings about it point back
/// into the original file. <see cref="JsonSchema"/> validates it; the YAML and TOML document models convert to it.
/// </summary>
/// <remarks>
/// Numbers keep their exact decimal value at any precision, so <c>0.1</c>, <c>multipleOf</c> and numbers beyond
/// <see cref="decimal"/> or <see cref="double"/> range compare exactly; only infinities and NaN are inexact. Objects
/// keep their properties in source order; a repeated name keeps its first position and its last value.
/// </remarks>
public sealed class JsonData {
    private static readonly JsonData[] NoItems = Array.Empty<JsonData>();
    private static readonly KeyValuePair<string, JsonData>[] NoProperties = Array.Empty<KeyValuePair<string, JsonData>>();
    private static readonly ReadOnlyCollection<JsonData> NoItemsView = new ReadOnlyCollection<JsonData>(NoItems);
    private static readonly ReadOnlyCollection<KeyValuePair<string, JsonData>> NoPropertiesView = new ReadOnlyCollection<KeyValuePair<string, JsonData>>(NoProperties);
    private const int LinearLookupLimit = 8;

    private readonly bool _boolean;
    private readonly string? _string;
    private readonly double _double;
    private readonly JsonExactNumber _exact;
    private readonly bool _hasExact;
    private readonly JsonData[] _items;
    private readonly KeyValuePair<string, JsonData>[] _properties;
    // Read-only views, so callers cannot cast Items or Properties back to the arrays and change them.
    private readonly ReadOnlyCollection<JsonData> _itemsView;
    private readonly ReadOnlyCollection<KeyValuePair<string, JsonData>> _propertiesView;
    private Dictionary<string, int>? _index;

    private JsonData(JsonDataKind kind, JsonDataLocation location) {
        Kind = kind;
        Line = location.Line;
        Column = location.Column;
        Source = location.Source;
        _items = NoItems;
        _properties = NoProperties;
        _itemsView = NoItemsView;
        _propertiesView = NoPropertiesView;
    }

    private JsonData(bool value, JsonDataLocation location) : this(JsonDataKind.Boolean, location) => _boolean = value;

    private JsonData(string value, JsonDataLocation location) : this(JsonDataKind.String, location) => _string = value;

    private JsonData(double value, JsonExactNumber exact, bool hasExact, JsonDataLocation location) : this(JsonDataKind.Number, location) {
        _double = value;
        _exact = exact;
        _hasExact = hasExact;
    }

    private JsonData(JsonData[] items, JsonDataLocation location) : this(JsonDataKind.Array, location) {
        _items = items;
        if (items.Length != 0) _itemsView = new ReadOnlyCollection<JsonData>(items);
    }

    private JsonData(KeyValuePair<string, JsonData>[] properties, JsonDataLocation location) : this(JsonDataKind.Object, location) {
        _properties = properties;
        if (properties.Length != 0) _propertiesView = new ReadOnlyCollection<KeyValuePair<string, JsonData>>(properties);
    }

    /// <summary>Gets the kind of value.</summary>
    public JsonDataKind Kind { get; }

    /// <summary>Gets the one-based line the value starts on, or 0 when it has no source position.</summary>
    public int Line { get; }

    /// <summary>Gets the one-based column the value starts at, or 0 when it has no source position.</summary>
    public int Column { get; }

    /// <summary>Gets the node the value was converted from, such as a <c>YamlNode</c> or <c>TomlNode</c>, or null.</summary>
    public object? Source { get; }

    /// <summary>Gets where the value came from.</summary>
    public JsonDataLocation Location => new JsonDataLocation(Line, Column, Source);

    /// <summary>Gets whether the value is a number with no fractional part, such as <c>3</c> or <c>3.0</c>.</summary>
    public bool IsInteger => Kind == JsonDataKind.Number && (_hasExact ? _exact.IsInteger : !double.IsInfinity(_double) && !double.IsNaN(_double) && Math.Floor(_double) == _double);

    /// <summary>Gets the items of an array, or an empty list for other kinds.</summary>
    public IReadOnlyList<JsonData> Items => _itemsView;

    /// <summary>Gets the properties of an object in source order, or an empty list for other kinds.</summary>
    public IReadOnlyList<KeyValuePair<string, JsonData>> Properties => _propertiesView;

    /// <summary>Creates the null value.</summary>
    public static JsonData CreateNull(JsonDataLocation location = default) => new JsonData(JsonDataKind.Null, location);

    /// <summary>Creates a boolean.</summary>
    public static JsonData CreateBoolean(bool value, JsonDataLocation location = default) => new JsonData(value, location);

    /// <summary>Creates a string.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    public static JsonData CreateString(string value, JsonDataLocation location = default) =>
        new JsonData(value ?? throw new ArgumentNullException(nameof(value)), location);

    /// <summary>Creates an exact number.</summary>
    public static JsonData CreateNumber(decimal value) => CreateNumber(value, default);

    /// <summary>Creates an exact number with its source location.</summary>
    public static JsonData CreateNumber(decimal value, JsonDataLocation location) {
        JsonExactNumber.TryParse(value.ToString(CultureInfo.InvariantCulture), out JsonExactNumber exact);
        return new JsonData((double)value, exact, true, location);
    }

    /// <summary>Creates an integer.</summary>
    public static JsonData CreateNumber(long value) => CreateNumber(value, default);

    /// <summary>Creates an integer with its source location.</summary>
    public static JsonData CreateNumber(long value, JsonDataLocation location) => new JsonData(value, JsonExactNumber.FromInt64(value), true, location);

    /// <summary>
    /// Creates a binary floating-point number. Finite values are kept as the exact decimal of their shortest round-trip
    /// form, so <c>0.1</c> stays 0.1; infinities and NaN, which JSON cannot write, are kept as inexact numbers.
    /// </summary>
    public static JsonData CreateNumber(double value) => CreateNumber(value, default);

    /// <summary>Creates a binary floating-point number with its source location; see <see cref="CreateNumber(double)"/>.</summary>
    public static JsonData CreateNumber(double value, JsonDataLocation location) {
        JsonExactNumber exact = default;
        bool hasExact = !double.IsNaN(value) && !double.IsInfinity(value) && JsonExactNumber.TryParse(value.ToString("R", CultureInfo.InvariantCulture), out exact);
        return new JsonData(value, hasExact ? exact : default, hasExact, location);
    }

    /// <summary>Creates a number from number text such as <c>-1.5e3</c>, keeping its exact value at any precision.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="text"/> is null.</exception>
    /// <exception cref="FormatException"><paramref name="text"/> is not a number.</exception>
    public static JsonData ParseNumber(string text, JsonDataLocation location = default) {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (!TryParseNumber(text, out double value, out JsonExactNumber exact, out bool hasExact)) {
            throw new FormatException("'" + text + "' is not a number, or its exponent is outside the supported range of nine digits.");
        }
        return new JsonData(value, exact, hasExact, location);
    }

    /// <summary>Creates an array from its items.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="items"/> or one of its items is null.</exception>
    public static JsonData CreateArray(IEnumerable<JsonData> items, JsonDataLocation location = default) {
        if (items is null) throw new ArgumentNullException(nameof(items));
        JsonData[] array = items.ToArray();
        foreach (JsonData item in array) {
            if (item is null) throw new ArgumentNullException(nameof(items), "Array items cannot be null; use JsonData.CreateNull.");
        }
        return new JsonData(array.Length == 0 ? NoItems : array, location);
    }

    /// <summary>Creates an object from its properties. A repeated name keeps its first position and its last value.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="properties"/>, a name or a value is null.</exception>
    public static JsonData CreateObject(IEnumerable<KeyValuePair<string, JsonData>> properties, JsonDataLocation location = default) {
        if (properties is null) throw new ArgumentNullException(nameof(properties));
        var list = new List<KeyValuePair<string, JsonData>>();
        Dictionary<string, int>? seen = null;
        foreach (KeyValuePair<string, JsonData> property in properties) {
            if (property.Key is null) throw new ArgumentNullException(nameof(properties), "Property names cannot be null.");
            if (property.Value is null) throw new ArgumentNullException(nameof(properties), "Property values cannot be null; use JsonData.CreateNull.");
            seen ??= new Dictionary<string, int>(StringComparer.Ordinal);
            if (seen.TryGetValue(property.Key, out int existing)) {
                list[existing] = property;
            } else {
                seen.Add(property.Key, list.Count);
                list.Add(property);
            }
        }
        return new JsonData(list.Count == 0 ? NoProperties : list.ToArray(), location);
    }

    /// <summary>Parses JSON text (RFC 8259), recording the line and column of every value.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="json"/> is null.</exception>
    /// <exception cref="FormatException">The text is not JSON, or nests deeper than 512 levels.</exception>
    public static JsonData Parse(string json) {
        if (json is null) throw new ArgumentNullException(nameof(json));
        return new JsonDataParser(json).ParseDocument();
    }

    /// <summary>Gets a boolean's value.</summary>
    /// <exception cref="InvalidOperationException">The value is not a boolean.</exception>
    public bool GetBoolean() => Kind == JsonDataKind.Boolean ? _boolean : throw WrongKind(JsonDataKind.Boolean);

    /// <summary>Gets a string's value.</summary>
    /// <exception cref="InvalidOperationException">The value is not a string.</exception>
    public string GetString() => Kind == JsonDataKind.String ? _string! : throw WrongKind(JsonDataKind.String);

    /// <summary>Gets a number as a <see cref="double"/>.</summary>
    /// <exception cref="InvalidOperationException">The value is not a number.</exception>
    public double GetDouble() => Kind == JsonDataKind.Number ? _double : throw WrongKind(JsonDataKind.Number);

    /// <summary>Gets a number as an exact <see cref="decimal"/>; returns false for non-numbers and numbers a decimal cannot hold exactly.</summary>
    public bool TryGetDecimal(out decimal value) {
        value = 0;
        return Kind == JsonDataKind.Number && _hasExact && _exact.TryToDecimal(out value);
    }

    /// <summary>Gets an object's property by name.</summary>
    public bool TryGetProperty(string name, out JsonData value) {
        if (name is null) throw new ArgumentNullException(nameof(name));
        int index = IndexOf(name);
        value = index < 0 ? null! : _properties[index].Value;
        return index >= 0;
    }

    /// <summary>Writes the value as compact JSON. Infinite and NaN numbers, which JSON cannot hold, are written as strings.</summary>
    public override string ToString() {
        var text = new StringBuilder();
        Write(text, this, int.MaxValue);
        return text.ToString();
    }

    /// <summary>Writes the value as compact JSON, cut to <paramref name="limit"/> characters with an ellipsis, without writing the rest.</summary>
    internal string ToString(int limit) {
        var text = new StringBuilder();
        if (Write(text, this, limit)) return text.ToString();
        text.Length = limit - 3;
        return text.Append("...").ToString();
    }

    /// <summary>Returns whether two values are equal in the JSON data model: numbers by value, objects regardless of property order.</summary>
    public static bool DeepEquals(JsonData? left, JsonData? right) => DeepEquals(left, right, null);

    // Values can share subtrees, as YAML aliases do; remembering container pairs already found equal keeps the
    // comparison linear in the number of distinct nodes.
    private static bool DeepEquals(JsonData? left, JsonData? right, HashSet<KeyValuePair<JsonData, JsonData>>? equal) {
        if (ReferenceEquals(left, right)) return true;
        if (left is null || right is null || left.Kind != right.Kind) return false;
        if (left.Kind is JsonDataKind.Array or JsonDataKind.Object) {
            equal ??= new HashSet<KeyValuePair<JsonData, JsonData>>(PairComparer.Instance);
            if (equal.Contains(new KeyValuePair<JsonData, JsonData>(left, right))) return true;
            if (!ContainersEqual(left, right, equal)) return false;
            equal.Add(new KeyValuePair<JsonData, JsonData>(left, right));
            return true;
        }
        switch (left.Kind) {
            case JsonDataKind.Null: return true;
            case JsonDataKind.Boolean: return left._boolean == right._boolean;
            case JsonDataKind.String: return string.Equals(left._string, right._string, StringComparison.Ordinal);
            default: return CompareNumbers(left, right) == 0;
        }
    }

    private static bool ContainersEqual(JsonData left, JsonData right, HashSet<KeyValuePair<JsonData, JsonData>> equal) {
        if (left.Kind == JsonDataKind.Array) {
            if (left._items.Length != right._items.Length) return false;
            for (int i = 0; i < left._items.Length; i++) {
                if (!DeepEquals(left._items[i], right._items[i], equal)) return false;
            }
            return true;
        }
        if (left._properties.Length != right._properties.Length) return false;
        foreach (KeyValuePair<string, JsonData> property in left._properties) {
            if (!right.TryGetProperty(property.Key, out JsonData other) || !DeepEquals(property.Value, other, equal)) return false;
        }
        return true;
    }

    /// <summary>
    /// A hash consistent with <see cref="DeepEquals(JsonData?, JsonData?)"/>. Pass the same <paramref name="hashes"/>
    /// for several values so shared subtrees are hashed once.
    /// </summary>
    internal static int DeepHash(JsonData value, Dictionary<JsonData, int> hashes) {
        if (value.Kind is not (JsonDataKind.Array or JsonDataKind.Object)) return ShallowHash(value);
        if (hashes.TryGetValue(value, out int known)) return known;
        int hash;
        unchecked {
            if (value.Kind == JsonDataKind.Array) {
                hash = 17;
                foreach (JsonData item in value._items) hash = hash * 31 + DeepHash(item, hashes);
            } else {
                hash = 19;
                // Order-independent combination, as property order does not matter.
                foreach (KeyValuePair<string, JsonData> property in value._properties) hash += StringComparer.Ordinal.GetHashCode(property.Key) ^ DeepHash(property.Value, hashes) * 7;
            }
        }
        hashes[value] = hash;
        return hash;
    }

    private static int ShallowHash(JsonData value) {
        unchecked {
            switch (value.Kind) {
                case JsonDataKind.Null: return 1;
                case JsonDataKind.Boolean: return value._boolean ? 3 : 2;
                case JsonDataKind.String: return StringComparer.Ordinal.GetHashCode(value._string!);
                default:
                    // Exact numbers are normalized, so equal values have equal fields.
                    return value._hasExact ? value._exact.Mantissa.GetHashCode() ^ value._exact.Exponent * 397 : value._double.GetHashCode();
            }
        }
    }

    /// <summary>Compares two numbers exactly when both are finite, otherwise as doubles.</summary>
    internal static int CompareNumbers(JsonData left, JsonData right) {
        if (left._hasExact && right._hasExact) return JsonExactNumber.Compare(left._exact, right._exact);
        if (left._hasExact != right._hasExact) {
            // A finite number is below positive infinity and above negative infinity and NaN, even when its double
            // approximation overflows.
            double other = left._hasExact ? right._double : left._double;
            int exactVersusOther = double.IsPositiveInfinity(other) ? -1 : 1;
            return left._hasExact ? exactVersusOther : -exactVersusOther;
        }
        return left._double.CompareTo(right._double);
    }

    internal bool HasExact => _hasExact;
    internal JsonExactNumber Exact => _exact;
    internal double DoubleValue => _double;

    internal int IndexOf(string name) {
        KeyValuePair<string, JsonData>[] properties = _properties;
        if (properties.Length <= LinearLookupLimit) {
            for (int i = 0; i < properties.Length; i++) {
                if (string.Equals(properties[i].Key, name, StringComparison.Ordinal)) return i;
            }
            return -1;
        }
        Dictionary<string, int>? index = _index;
        if (index is null) {
            index = new Dictionary<string, int>(properties.Length, StringComparer.Ordinal);
            for (int i = 0; i < properties.Length; i++) index[properties[i].Key] = i;
            _index = index;
        }
        return index.TryGetValue(name, out int found) ? found : -1;
    }

    internal static bool TryParseNumber(string text, out double value, out JsonExactNumber exact, out bool hasExact) {
        hasExact = JsonExactNumber.TryParse(text, out exact);
        bool parsed = double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        if (hasExact) {
            // .NET Framework reports overflow instead of returning infinity.
            if (!parsed) value = exact.Mantissa.Sign < 0 ? double.NegativeInfinity : double.PositiveInfinity;
            return true;
        }
        // Number text the exact parser rejects has an exponent beyond its range; a double would silently turn it into
        // infinity or zero. Only the words for infinity and NaN stay inexact.
        if (!parsed || !(double.IsNaN(value) || double.IsInfinity(value))) return false;
        foreach (char c in text) {
            if (c >= '0' && c <= '9') return false;
        }
        return true;
    }

    private InvalidOperationException WrongKind(JsonDataKind expected) =>
        new InvalidOperationException("The value is " + Describe(Kind) + ", not " + Describe(expected) + ".");

    internal static string Describe(JsonDataKind kind) => kind switch {
        JsonDataKind.Null => "null",
        JsonDataKind.Boolean => "a boolean",
        JsonDataKind.Number => "a number",
        JsonDataKind.String => "a string",
        JsonDataKind.Array => "an array",
        _ => "an object"
    };

    /// <summary>Appends the value as compact JSON; stops and returns false once the text exceeds <paramref name="limit"/> characters.</summary>
    internal static bool Write(StringBuilder text, JsonData value, int limit) {
        if (text.Length > limit) return false;
        switch (value.Kind) {
            case JsonDataKind.Null: text.Append("null"); break;
            case JsonDataKind.Boolean: text.Append(value._boolean ? "true" : "false"); break;
            case JsonDataKind.String: WriteString(text, value._string!, limit); break;
            case JsonDataKind.Number:
                if (value._hasExact) text.Append(value._exact.ToString());
                else if (double.IsNaN(value._double) || double.IsInfinity(value._double)) WriteString(text, value._double.ToString(CultureInfo.InvariantCulture), limit);
                else text.Append(value._double.ToString("R", CultureInfo.InvariantCulture));
                break;
            case JsonDataKind.Array:
                text.Append('[');
                for (int i = 0; i < value._items.Length; i++) {
                    if (i > 0) text.Append(',');
                    if (!Write(text, value._items[i], limit)) return false;
                }
                text.Append(']');
                break;
            default:
                text.Append('{');
                for (int i = 0; i < value._properties.Length; i++) {
                    if (i > 0) text.Append(',');
                    if (!WriteString(text, value._properties[i].Key, limit)) return false;
                    text.Append(':');
                    if (!Write(text, value._properties[i].Value, limit)) return false;
                }
                text.Append('}');
                break;
        }
        return text.Length <= limit;
    }

    /// <summary>Appends a JSON string; stops and returns false once the text exceeds <paramref name="limit"/> characters.</summary>
    internal static bool WriteString(StringBuilder text, string value, int limit) {
        text.Append('"');
        foreach (char c in value) {
            if (text.Length > limit) return false;
            switch (c) {
                case '"': text.Append("\\\""); break;
                case '\\': text.Append("\\\\"); break;
                case '\n': text.Append("\\n"); break;
                case '\r': text.Append("\\r"); break;
                case '\t': text.Append("\\t"); break;
                case '\b': text.Append("\\b"); break;
                case '\f': text.Append("\\f"); break;
                default:
                    if (c < 0x20) text.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                    else text.Append(c);
                    break;
            }
        }
        text.Append('"');
        return text.Length <= limit;
    }
}

/// <summary>Where a <see cref="JsonData"/> value came from: its one-based line and column, and the source node.</summary>
public readonly struct JsonDataLocation {
    /// <summary>Creates a location. Lines and columns below 1 mean "unknown" and are stored as 0.</summary>
    public JsonDataLocation(int line, int column, object? source = null) {
        Line = line < 1 ? 0 : line;
        Column = column < 1 ? 0 : column;
        Source = source;
    }

    /// <summary>Gets the one-based line, or 0 when unknown.</summary>
    public int Line { get; }

    /// <summary>Gets the one-based column, or 0 when unknown.</summary>
    public int Column { get; }

    /// <summary>Gets the source node, such as a <c>YamlNode</c> or <c>TomlNode</c>, or null.</summary>
    public object? Source { get; }
}

/// <summary>Compares pairs of <see cref="JsonData"/> instances by reference.</summary>
internal sealed class PairComparer : IEqualityComparer<KeyValuePair<JsonData, JsonData>> {
    public static readonly PairComparer Instance = new PairComparer();
    public bool Equals(KeyValuePair<JsonData, JsonData> x, KeyValuePair<JsonData, JsonData> y) => ReferenceEquals(x.Key, y.Key) && ReferenceEquals(x.Value, y.Value);
    public int GetHashCode(KeyValuePair<JsonData, JsonData> obj) => unchecked(RuntimeHelpers.GetHashCode(obj.Key) * 397 ^ RuntimeHelpers.GetHashCode(obj.Value));
}

/// <summary>Compares <see cref="JsonData"/> instances by reference, for caches keyed by schema node.</summary>
internal sealed class JsonDataReferenceComparer : IEqualityComparer<JsonData> {
    public static readonly JsonDataReferenceComparer Instance = new JsonDataReferenceComparer();
    public bool Equals(JsonData? x, JsonData? y) => ReferenceEquals(x, y);
    public int GetHashCode(JsonData obj) => RuntimeHelpers.GetHashCode(obj);
}
