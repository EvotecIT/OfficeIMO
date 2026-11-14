using System.Numerics;

namespace OfficeIMO.Toml;

/// <summary>
/// Typed reads of loaded values straight from their source text, without the boxed <see cref="TomlValue.Value"/>.
/// The rules match <see cref="TomlNode.GetValue{T}"/>: integers read as floating types, and a float reads as an integer
/// type only when it is whole and in range.
/// </summary>
internal static class TomlValueText {
    public static string GetString(TomlReadNode value) =>
        value.TryGetRaw(out ReadOnlySpan<char> raw) ? TomlScalars.Decode(raw, value.StringStyle) : (string)value.DecodedValue;

    public static long GetInt64(TomlReadNode value) {
        if (value.TryGetRaw(out ReadOnlySpan<char> raw)) {
            TomlScalars.TryParseInteger(raw, out long number, out _, out _);
            return number;
        }
        return (long)value.DecodedValue;
    }

    public static bool GetBoolean(TomlReadNode value) =>
        value.TryGetRaw(out ReadOnlySpan<char> raw) ? raw[0] == 't' : (bool)value.DecodedValue;

    public static double GetDouble(TomlReadNode value) =>
        value.TryGetRaw(out ReadOnlySpan<char> raw) ? TomlScalars.ParseFloat(raw) : (double)value.DecodedValue;

    public static TomlDateTime GetDateTime(TomlReadNode value) {
        if (value.TryGetRaw(out ReadOnlySpan<char> raw)) {
            TomlScalars.TryParseDateTime(raw, out TomlDateTime dateTime, out _);
            return dateTime;
        }
        return (TomlDateTime)value.DecodedValue;
    }

    /// <summary>Reads an integer, or a whole float in range, as a 64-bit integer.</summary>
    public static bool TryGetWhole(TomlReadNode value, out long number) {
        if (value.Kind == TomlNodeKind.Integer) {
            number = GetInt64(value);
            return true;
        }
        if (value.Kind == TomlNodeKind.Float) {
            double real = GetDouble(value);
            if (!double.IsNaN(real) && !double.IsInfinity(real) && Math.Floor(real) == real && real >= long.MinValue && real < long.MaxValue) {
                number = (long)real;
                return true;
            }
        }
        number = 0;
        return false;
    }

    public static bool TryGetReal(TomlReadNode value, out double number) {
        if (value.Kind == TomlNodeKind.Float) {
            number = GetDouble(value);
            return true;
        }
        if (value.Kind == TomlNodeKind.Integer) {
            number = GetInt64(value);
            return true;
        }
        number = 0;
        return false;
    }

    /// <summary>A float's text without digit separators, for an exact decimal read.</summary>
    public static string FloatText(TomlReadNode value) {
        if (value.TryGetRaw(out ReadOnlySpan<char> raw)) return raw.ToString().Replace("_", string.Empty);
        return ((double)value.DecodedValue).ToString("R", CultureInfo.InvariantCulture);
    }
}

/// <summary>A type written as one TOML value, read only from the value kinds that can hold it.</summary>
internal abstract class TomlScalarConverter<T> : TomlConverter<T> {
    public override T Read(TomlReadNode node, TomlReadState state) {
        if (node.IsValue && TryRead(node, out T result)) return result;
        throw TomlSerializationErrors.TypeMismatch(node, typeof(T), Detail(node));
    }

    /// <summary>Why a value of the right kind does not fit, such as a range, or null.</summary>
    protected virtual string? Detail(TomlReadNode node) => null;

    protected abstract bool TryRead(TomlReadNode value, out T result);

    /// <summary>Formats a value as a key; keys are strings, so the value's own text is used.</summary>
    public override string? WriteKey(T value) => FormatText(value);

    public override bool TryReadKey(string key, out T value) => TryParseText(key, out value);

    protected virtual string FormatText(T value) => Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;

    protected virtual bool TryParseText(string text, out T value) {
        value = default!;
        return false;
    }
}

internal sealed class TomlStringConverter : TomlScalarConverter<string> {
    public static readonly TomlStringConverter Instance = new TomlStringConverter();

    protected override bool TryRead(TomlReadNode value, out string result) {
        if (value.Kind != TomlNodeKind.String) {
            result = null!;
            return false;
        }
        result = TomlValueText.GetString(value);
        return true;
    }

    public override void Write(TomlEmitter emitter, string value, TomlWriteState state) => emitter.WriteString(value);

    protected override string FormatText(string value) => value;

    protected override bool TryParseText(string text, out string value) {
        value = text;
        return true;
    }
}

internal sealed class TomlBooleanConverter : TomlScalarConverter<bool> {
    public static readonly TomlBooleanConverter Instance = new TomlBooleanConverter();

    protected override bool TryRead(TomlReadNode value, out bool result) {
        result = value.Kind == TomlNodeKind.Boolean && TomlValueText.GetBoolean(value);
        return value.Kind == TomlNodeKind.Boolean;
    }

    public override void Write(TomlEmitter emitter, bool value, TomlWriteState state) => emitter.WriteBoolean(value);

    protected override string FormatText(bool value) => value ? "true" : "false";

    protected override bool TryParseText(string text, out bool value) {
        value = text == "true";
        return value || text == "false";
    }
}

/// <summary>The integer types: read from integers and whole floats in the type's range, written as integers.</summary>
internal sealed class TomlIntegerConverter<T> : TomlScalarConverter<T> where T : struct {
    private readonly long _min;
    private readonly long _max;
    private readonly Func<long, T> _from;
    private readonly Func<T, long> _to;

    public TomlIntegerConverter(long min, long max, Func<long, T> from, Func<T, long> to) {
        _min = min;
        _max = max;
        _from = from;
        _to = to;
    }

    protected override bool TryRead(TomlReadNode value, out T result) {
        if (TomlValueText.TryGetWhole(value, out long number) && number >= _min && number <= _max) {
            result = _from(number);
            return true;
        }
        result = default;
        return false;
    }

    protected override string? Detail(TomlReadNode node) =>
        node.IsValue && TomlValueText.TryGetWhole(node, out _) ? "the value is outside the range of " + typeof(T).Name : null;

    public override void Write(TomlEmitter emitter, T value, TomlWriteState state) => emitter.WriteInteger(_to(value));

    protected override bool TryParseText(string text, out T value) {
        if (long.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out long number) && number >= _min && number <= _max) {
            value = _from(number);
            return true;
        }
        value = default;
        return false;
    }
}

internal sealed class TomlUInt64Converter : TomlScalarConverter<ulong> {
    public static readonly TomlUInt64Converter Instance = new TomlUInt64Converter();

    protected override bool TryRead(TomlReadNode value, out ulong result) {
        if (TomlValueText.TryGetWhole(value, out long number) && number >= 0) {
            result = (ulong)number;
            return true;
        }
        result = 0;
        return false;
    }

    public override void Write(TomlEmitter emitter, ulong value, TomlWriteState state) {
        if (value > long.MaxValue) throw TomlSerializationErrors.OutOfRange(value.ToString(CultureInfo.InvariantCulture));
        emitter.WriteInteger((long)value);
    }

    protected override bool TryParseText(string text, out ulong value) => ulong.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value);
}

internal sealed class TomlBigIntegerConverter : TomlScalarConverter<BigInteger> {
    public static readonly TomlBigIntegerConverter Instance = new TomlBigIntegerConverter();

    protected override bool TryRead(TomlReadNode value, out BigInteger result) {
        bool whole = TomlValueText.TryGetWhole(value, out long number);
        result = number;
        return whole;
    }

    public override void Write(TomlEmitter emitter, BigInteger value, TomlWriteState state) {
        if (value < long.MinValue || value > long.MaxValue) throw TomlSerializationErrors.OutOfRange(value.ToString(CultureInfo.InvariantCulture));
        emitter.WriteInteger((long)value);
    }

    protected override bool TryParseText(string text, out BigInteger value) => BigInteger.TryParse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
}

internal sealed class TomlDoubleConverter : TomlScalarConverter<double> {
    public static readonly TomlDoubleConverter Instance = new TomlDoubleConverter();

    protected override bool TryRead(TomlReadNode value, out double result) => TomlValueText.TryGetReal(value, out result);

    public override void Write(TomlEmitter emitter, double value, TomlWriteState state) => emitter.WriteFloat(value);

    protected override string FormatText(double value) => value.ToString("R", CultureInfo.InvariantCulture);

    protected override bool TryParseText(string text, out double value) => double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
}

/// <summary>Singles are written with their own shortest text, so <c>1.1f</c> stays <c>1.1</c> instead of its double expansion.</summary>
internal sealed class TomlSingleConverter : TomlScalarConverter<float> {
    public static readonly TomlSingleConverter Instance = new TomlSingleConverter();

    protected override bool TryRead(TomlReadNode value, out float result) {
        bool read = TomlValueText.TryGetReal(value, out double number);
        result = (float)number;
        return read;
    }

    public override void Write(TomlEmitter emitter, float value, TomlWriteState state) {
        if (float.IsNaN(value) || float.IsInfinity(value)) {
            emitter.WriteFloat(value);
            return;
        }
        emitter.WriteFloatText(value.ToString("R", CultureInfo.InvariantCulture));
    }

    protected override string FormatText(float value) => value.ToString("R", CultureInfo.InvariantCulture);

    protected override bool TryParseText(string text, out float value) => float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
}

/// <summary>Decimals are written with their exact digits and read from a float's text, so no digit goes through a double.</summary>
internal sealed class TomlDecimalConverter : TomlScalarConverter<decimal> {
    public static readonly TomlDecimalConverter Instance = new TomlDecimalConverter();

    protected override bool TryRead(TomlReadNode value, out decimal result) {
        if (value.Kind == TomlNodeKind.Integer) {
            result = TomlValueText.GetInt64(value);
            return true;
        }
        result = 0;
        return value.Kind == TomlNodeKind.Float
            && decimal.TryParse(TomlValueText.FloatText(value), NumberStyles.Float, CultureInfo.InvariantCulture, out result);
    }

    public override void Write(TomlEmitter emitter, decimal value, TomlWriteState state) => emitter.WriteFloatText(value.ToString(CultureInfo.InvariantCulture));

    protected override bool TryParseText(string text, out decimal value) => decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
}

internal sealed class TomlCharConverter : TomlScalarConverter<char> {
    public static readonly TomlCharConverter Instance = new TomlCharConverter();

    protected override bool TryRead(TomlReadNode value, out char result) {
        string? text = value.Kind == TomlNodeKind.String ? TomlValueText.GetString(value) : null;
        result = text is { Length: 1 } ? text[0] : '\0';
        return text is { Length: 1 };
    }

    public override void Write(TomlEmitter emitter, char value, TomlWriteState state) => emitter.WriteString(value.ToString());

    protected override string FormatText(char value) => value.ToString();

    protected override bool TryParseText(string text, out char value) {
        value = text.Length == 1 ? text[0] : '\0';
        return text.Length == 1;
    }
}

/// <summary>Converts a source value to <typeparamref name="T"/>, or fails.</summary>
internal delegate bool TomlTryConvert<TSource, T>(TSource source, out T value);

/// <summary>The types written as strings.</summary>
internal static class TomlTextConverters {
    public static readonly TomlTextConverter<Guid> GuidText = new TomlTextConverter<Guid>(Guid.TryParse, static id => id.ToString());

    public static readonly TomlTextConverter<Uri> UriText = new TomlTextConverter<Uri>(
        static (string text, out Uri uri) => Uri.TryCreate(text, UriKind.RelativeOrAbsolute, out uri!), static uri => uri.OriginalString);

    public static readonly TomlTextConverter<byte[]> Base64 = new TomlTextConverter<byte[]>(static (string text, out byte[] bytes) => {
        try {
            bytes = Convert.FromBase64String(text);
            return true;
        } catch (FormatException) {
            bytes = null!;
            return false;
        }
    }, static bytes => Convert.ToBase64String(bytes));
}

/// <summary>A type written as a string and parsed from it: <see cref="Guid"/>, <see cref="Uri"/> and <see cref="byte"/> arrays (Base64).</summary>
internal sealed class TomlTextConverter<T> : TomlScalarConverter<T> {
    private readonly TomlTryConvert<string, T> _parse;
    private readonly Func<T, string> _format;

    internal TomlTextConverter(TomlTryConvert<string, T> parse, Func<T, string> format) {
        _parse = parse;
        _format = format;
    }

    protected override bool TryRead(TomlReadNode value, out T result) {
        if (value.Kind != TomlNodeKind.String) {
            result = default!;
            return false;
        }
        return _parse(TomlValueText.GetString(value), out result);
    }

    public override void Write(TomlEmitter emitter, T value, TomlWriteState state) => emitter.WriteString(_format(value));

    protected override string FormatText(T value) => _format(value);

    protected override bool TryParseText(string text, out T value) => _parse(text, out value);
}

/// <summary>
/// The date and time types, mapped to the TOML kinds: <see cref="DateTimeOffset"/> to an offset date-time,
/// <see cref="DateTime"/> to a local date-time (or an offset date-time for UTC and local kinds), <c>DateOnly</c> to a
/// local date, <c>TimeOnly</c> to a local time and <see cref="TomlDateTime"/> to any kind unchanged. Reading follows
/// <see cref="TomlNode.GetValue{T}"/>.
/// </summary>
internal sealed class TomlDateTimeConverter<T> : TomlScalarConverter<T> {
    private readonly TomlTryConvert<TomlDateTime, T> _read;
    private readonly Func<T, TomlDateTime> _write;

    internal TomlDateTimeConverter(TomlTryConvert<TomlDateTime, T> read, Func<T, TomlDateTime> write) {
        _read = read;
        _write = write;
    }

    protected override bool TryRead(TomlReadNode value, out T result) {
        if (value.Kind != TomlNodeKind.DateTime) {
            result = default!;
            return false;
        }
        return _read(TomlValueText.GetDateTime(value), out result);
    }

    public override void Write(TomlEmitter emitter, T value, TomlWriteState state) => emitter.WriteDateTime(_write(value));

    protected override string FormatText(T value) => _write(value).ToString();

    protected override bool TryParseText(string text, out T value) {
        if (TomlScalars.TryParseDateTime(text.AsSpan(), out TomlDateTime parsed, out _) is not null) {
            value = default!;
            return false;
        }
        return _read(parsed, out value);
    }
}

/// <summary>The date and time types.</summary>
internal static class TomlDateTimeConverters {
    public static readonly TomlDateTimeConverter<TomlDateTime> Toml = new TomlDateTimeConverter<TomlDateTime>(
        static (TomlDateTime source, out TomlDateTime value) => {
            value = source;
            return true;
        }, static v => v);

    public static readonly TomlDateTimeConverter<DateTimeOffset> Offset = new TomlDateTimeConverter<DateTimeOffset>(
        static (TomlDateTime source, out DateTimeOffset value) => {
            bool offset = source.Kind == TomlDateTimeKind.OffsetDateTime;
            value = offset ? source.ToDateTimeOffset() : default;
            return offset;
        }, TomlDateTime.FromDateTimeOffset);

    public static readonly TomlDateTimeConverter<DateTime> Local = new TomlDateTimeConverter<DateTime>(
        static (TomlDateTime source, out DateTime value) => {
            bool dated = source.Kind != TomlDateTimeKind.LocalTime;
            value = dated ? source.ToDateTime() : default;
            return dated;
        }, TomlDateTime.FromDateTime);

#if NET6_0_OR_GREATER
    public static readonly TomlDateTimeConverter<DateOnly> Date = new TomlDateTimeConverter<DateOnly>(
        static (TomlDateTime source, out DateOnly value) => {
            bool dated = source.Kind != TomlDateTimeKind.LocalTime;
            value = dated ? new DateOnly(source.Year, source.Month, source.Day) : default;
            return dated;
        }, static v => TomlDateTime.FromDate(v.Year, v.Month, v.Day));

    public static readonly TomlDateTimeConverter<TimeOnly> Time = new TomlDateTimeConverter<TimeOnly>(
        static (TomlDateTime source, out TimeOnly value) => {
            bool timed = source.Kind != TomlDateTimeKind.LocalDate && source.Second != 60;
            value = timed ? new TimeOnly(source.ToTimeOfDay().Ticks) : default;
            return timed;
        }, static v => TomlDateTime.FromTime(v.Hour, v.Minute, v.Second, (int)(v.Ticks % TimeSpan.TicksPerSecond) * 100));
#endif
}

/// <summary>
/// <see cref="TimeSpan"/> as a local time when it is a time of day (zero up to one day), as TOML has no duration;
/// any other span is written as its invariant text (<c>1.02:03:04</c>). Reading accepts both.
/// </summary>
internal sealed class TomlTimeSpanConverter : TomlScalarConverter<TimeSpan> {
    public static readonly TomlTimeSpanConverter Instance = new TomlTimeSpanConverter();

    protected override bool TryRead(TomlReadNode value, out TimeSpan result) {
        result = default;
        if (value.Kind == TomlNodeKind.String) return TimeSpan.TryParse(TomlValueText.GetString(value), CultureInfo.InvariantCulture, out result);
        if (value.Kind != TomlNodeKind.DateTime) return false;
        TomlDateTime time = TomlValueText.GetDateTime(value);
        if (time.Kind != TomlDateTimeKind.LocalTime) return false;
        result = time.ToTimeOfDay();
        return true;
    }

    public override void Write(TomlEmitter emitter, TimeSpan value, TomlWriteState state) {
        if (value >= TimeSpan.Zero && value < TimeSpan.FromDays(1)) {
            emitter.WriteDateTime(TomlDateTime.FromTime(value.Hours, value.Minutes, value.Seconds, (int)(value.Ticks % TimeSpan.TicksPerSecond) * 100));
        } else {
            emitter.WriteString(value.ToString("c", CultureInfo.InvariantCulture));
        }
    }

    protected override string FormatText(TimeSpan value) => value.ToString("c", CultureInfo.InvariantCulture);

    protected override bool TryParseText(string text, out TimeSpan value) => TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out value);
}

/// <summary>A nullable value type: TOML has no null, so a present value is read and written as the underlying type.</summary>
internal sealed class TomlNullableConverter<T> : TomlConverter<T?> where T : struct {
    private TomlTypeInfo<T>? _underlying;

    public override TomlShape Shape => _underlying!.StaticShape;

    public override bool HasDynamicShape => _underlying!.Converter.HasDynamicShape;

    public override bool IsDictionary => false;

    public override TomlShape GetShape(T? value) => _underlying!.GetShape(value!.Value);

    public override void Configure(TomlSerializerOptions options) => _underlying = (TomlTypeInfo<T>)options.GetTypeInfo(typeof(T));

    public override T? Read(TomlReadNode node, TomlReadState state) => _underlying!.Read(node, state);

    public override void Write(TomlEmitter emitter, T? value, TomlWriteState state) => _underlying!.Write(emitter, value!.Value, state);

    public override void WriteBody(TomlEmitter emitter, T? value, TomlWriteState state) => _underlying!.WriteBody(emitter, value!.Value, state);

    public override string? WriteKey(T? value) => _underlying!.Converter.WriteKey(value!.Value);

    public override bool TryReadKey(string key, out T? value) {
        bool read = _underlying!.Converter.TryReadKey(key, out T underlying);
        value = read ? underlying : null;
        return read;
    }
}
