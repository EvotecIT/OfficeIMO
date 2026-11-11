namespace OfficeIMO.Toml;

/// <summary>Converts between TOML values and .NET values, for typed reads and for values passed to editing methods.</summary>
internal static class TomlValueConverter {
    private const int MaxNesting = 256;

    public static bool IsScalar(object value) => TryCreateScalar(value, out _, out _);

    /// <summary>Normalizes a .NET scalar to the stored form: string, long, double, bool or <see cref="TomlDateTime"/>.</summary>
    public static bool TryCreateScalar(object value, out TomlNodeKind kind, out object normalized) {
        switch (value) {
            case string text:
                kind = TomlNodeKind.String;
                normalized = text;
                return true;
            case bool flag:
                kind = TomlNodeKind.Boolean;
                normalized = flag;
                return true;
            case int or long or short or sbyte or byte or ushort or uint:
                kind = TomlNodeKind.Integer;
                normalized = Convert.ToInt64(value, CultureInfo.InvariantCulture);
                return true;
            case ulong unsigned:
                if (unsigned > long.MaxValue) throw new ArgumentOutOfRangeException(nameof(value), value, "TOML integers are 64-bit signed.");
                kind = TomlNodeKind.Integer;
                normalized = (long)unsigned;
                return true;
            case double or float or decimal:
                kind = TomlNodeKind.Float;
                normalized = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            case TomlDateTime dateTime:
                kind = TomlNodeKind.DateTime;
                normalized = dateTime;
                return true;
            case DateTimeOffset offset:
                kind = TomlNodeKind.DateTime;
                normalized = TomlDateTime.FromDateTimeOffset(offset);
                return true;
            case DateTime local:
                kind = TomlNodeKind.DateTime;
                normalized = TomlDateTime.FromDateTime(local);
                return true;
#if NET6_0_OR_GREATER
            case DateOnly date:
                kind = TomlNodeKind.DateTime;
                normalized = TomlDateTime.FromDate(date.Year, date.Month, date.Day);
                return true;
            case TimeOnly time:
                kind = TomlNodeKind.DateTime;
                normalized = TomlDateTime.FromTime(time.Hour, time.Minute, time.Second, (int)(time.Ticks % TimeSpan.TicksPerSecond) * 100);
                return true;
#endif
            default:
                kind = default;
                normalized = null!;
                return false;
        }
    }

    /// <summary>Creates a new node for a value passed to an editing method.</summary>
    public static TomlNode CreateNode(TomlDocument document, object value, string parameter, TomlStringStyle preferred) =>
        CreateNode(document, value, parameter, preferred, 0);

    private static TomlNode CreateNode(TomlDocument document, object value, string parameter, TomlStringStyle preferred, int nesting) {
        if (value is null) throw new ArgumentNullException(parameter, "TOML has no null value.");
        if (nesting > MaxNesting) throw new ArgumentException("The value nests more than " + MaxNesting + " levels deep or contains itself.", parameter);
        if (value is TomlNode) throw new ArgumentException("Nodes cannot be moved or shared between tables; pass the value instead.", parameter);
        if (TryCreateScalar(value, out TomlNodeKind kind, out object normalized)) {
            var scalar = new TomlValue(document, kind) { _value = normalized };
            if (kind == TomlNodeKind.String) scalar._style = TomlFormatter.FitStyle((string)normalized, preferred);
            return scalar;
        }
        if (value is IDictionary dictionary) {
            var table = new TomlTable(document, TomlTableKind.Inline);
            foreach (DictionaryEntry entry in dictionary) {
                if (entry.Key is not string key) throw new ArgumentException("Dictionary keys must be strings.", parameter);
                if (table.IndexOf(key) >= 0) throw new ArgumentException("The dictionary has the key '" + key + "' twice.", parameter);
                table.AddEntry(key, CreateNode(document, entry.Value!, parameter, preferred, nesting + 1));
            }
            return table;
        }
        if (value is IEnumerable items) {
            var array = new TomlArray(document);
            foreach (object? item in items) array.AddLoaded(CreateNode(document, item!, parameter, preferred, nesting + 1));
            return array;
        }
        throw new ArgumentException("A " + value.GetType().Name + " is not a TOML value; use a string, number, boolean, date-time, sequence or dictionary.", parameter);
    }

    public static bool ScalarEquals(object current, object candidate) => current switch {
        string text => candidate is string other && string.Equals(text, other, StringComparison.Ordinal),
        long number => candidate is long other && number == other,
        double number => candidate is double other && (number.Equals(other) && (number != 0 || BitConverter.DoubleToInt64Bits(number) == BitConverter.DoubleToInt64Bits(other))),
        bool flag => candidate is bool other && flag == other,
        TomlDateTime dateTime => candidate is TomlDateTime other && dateTime == other,
        _ => false
    };

    public static void RequireSupported<T>() {
        Type type = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
        if (type == typeof(string) || type == typeof(bool) || type == typeof(int) || type == typeof(long) || type == typeof(double)
            || type == typeof(float) || type == typeof(decimal) || type == typeof(TomlDateTime) || type == typeof(DateTime)
            || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) || type == typeof(object)) {
            return;
        }
#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly) || type == typeof(TimeOnly)) return;
#endif
        throw new NotSupportedException("TOML values cannot be read as " + typeof(T).Name + ".");
    }

    public static bool TryConvert<T>(TomlValue node, out T result) {
        RequireSupported<T>();
        Type type = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
        object value = node.Value;
        object? converted = null;
        if (type == typeof(object)) {
            converted = value;
        } else if (type == typeof(string)) {
            converted = node.GetText();
        } else if (type == typeof(bool)) {
            if (value is bool flag) converted = flag;
        } else if (type == typeof(long) || type == typeof(int)) {
            long? whole = value switch {
                long number => number,
                double number when IsWhole(number, long.MinValue, long.MaxValue) => (long)number,
                _ => null
            };
            if (whole is long w) {
                if (type == typeof(long)) converted = w;
                else if (w >= int.MinValue && w <= int.MaxValue) converted = (int)w;
            }
        } else if (type == typeof(double)) {
            if (value is double number) converted = number;
            else if (value is long integer) converted = (double)integer;
        } else if (type == typeof(float)) {
            if (value is double number) converted = (float)number;
            else if (value is long integer) converted = (float)integer;
        } else if (type == typeof(decimal)) {
            if (value is long integer) converted = (decimal)integer;
            else if (value is double number && !double.IsNaN(number) && !double.IsInfinity(number) && Math.Abs(number) < 7.9e28) converted = (decimal)number;
        } else if (value is TomlDateTime dateTime) {
            converted = ConvertDateTime(dateTime, type);
        }
        if (converted is null) {
            result = default!;
            return false;
        }
        result = (T)converted;
        return true;
    }

    private static object? ConvertDateTime(TomlDateTime value, Type type) {
        if (type == typeof(TomlDateTime)) return value;
        if (type == typeof(DateTimeOffset)) return value.Kind == TomlDateTimeKind.OffsetDateTime ? value.ToDateTimeOffset() : null;
        if (type == typeof(DateTime)) return value.Kind == TomlDateTimeKind.LocalTime ? null : value.ToDateTime();
        if (type == typeof(TimeSpan)) return value.Kind == TomlDateTimeKind.LocalDate ? null : value.ToTimeOfDay();
#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly)) return value.Kind == TomlDateTimeKind.LocalTime ? null : new DateOnly(value.Year, value.Month, value.Day);
        if (type == typeof(TimeOnly)) return value.Kind == TomlDateTimeKind.LocalDate || value.Second == 60 ? null : new TimeOnly(value.ToTimeOfDay().Ticks);
#endif
        return null;
    }

    private static bool IsWhole(double value, double min, double max) =>
        !double.IsNaN(value) && !double.IsInfinity(value) && Math.Floor(value) == value && value >= min && value < max;
}
