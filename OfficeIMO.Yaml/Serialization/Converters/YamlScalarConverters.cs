using System.Numerics;

namespace OfficeIMO.Yaml;

/// <summary>
/// Shared scalar reading for the built-in converters. Scalars are read by the target type: the decoded text is parsed
/// with the document's schema rules whatever its quoting, so <c>"42"</c> reads as an integer and <c>'true'</c> as a
/// boolean when the target asks for one. Null handling happens before a converter runs.
/// </summary>
internal static class YamlScalarText {
    public static ReadOnlySpan<char> Read(YamlEventCursor cursor, YamlReadState state, Type type) {
        if (cursor.TokenType != YamlTokenType.Scalar) throw YamlSerializationErrors.TypeMismatch(cursor, state, type);
        return cursor.ScalarSpan;
    }

    public static long ReadInteger(YamlEventCursor cursor, YamlReadState state, Type type, long min, long max) {
        ReadOnlySpan<char> text = Read(cursor, state, type);
        if (!YamlSchemaResolver.TryParseInt64(text, cursor.Schema, out long value)) {
            throw YamlSerializationErrors.TypeMismatch(cursor, state, type, "not an integer, or outside the range of a 64-bit integer");
        }
        if (value < min || value > max) throw YamlSerializationErrors.TypeMismatch(cursor, state, type, "outside the range of the type");
        return value;
    }

    public static double ReadDouble(YamlEventCursor cursor, YamlReadState state, Type type) {
        ReadOnlySpan<char> text = Read(cursor, state, type);
        if (!YamlSchemaResolver.TryParseDouble(text, cursor.Schema, out double value)) {
            throw YamlSerializationErrors.TypeMismatch(cursor, state, type, "not a number");
        }
        return value;
    }

    public static void WriteInvariant(YamlWriter writer, IFormattable value, string? format) =>
        writer.WritePlainValue(value.ToString(format, CultureInfo.InvariantCulture).AsSpan());
}

internal sealed class YamlStringConverter : YamlConverter<string> {
    public static YamlStringConverter Instance { get; } = new YamlStringConverter();

    public override string Read(YamlEventCursor cursor, YamlReadState state) {
        if (cursor.TokenType != YamlTokenType.Scalar) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(string));
        return cursor.GetString();
    }

    public override void Write(YamlWriter writer, string value, YamlWriteState state) => writer.WriteScalar(value);
}

internal sealed class YamlBooleanConverter : YamlConverter<bool> {
    public static YamlBooleanConverter Instance { get; } = new YamlBooleanConverter();

    public override bool Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(bool));
        if (!YamlSchemaResolver.TryParseBoolean(text, cursor.Schema, out bool value)) {
            throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(bool));
        }
        return value;
    }

    public override void Write(YamlWriter writer, bool value, YamlWriteState state) => writer.WriteScalar(value);
}

internal sealed class YamlByteConverter : YamlConverter<byte> {
    public static YamlByteConverter Instance { get; } = new YamlByteConverter();
    public override byte Read(YamlEventCursor cursor, YamlReadState state) => (byte)YamlScalarText.ReadInteger(cursor, state, typeof(byte), byte.MinValue, byte.MaxValue);
    public override void Write(YamlWriter writer, byte value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlSByteConverter : YamlConverter<sbyte> {
    public static YamlSByteConverter Instance { get; } = new YamlSByteConverter();
    public override sbyte Read(YamlEventCursor cursor, YamlReadState state) => (sbyte)YamlScalarText.ReadInteger(cursor, state, typeof(sbyte), sbyte.MinValue, sbyte.MaxValue);
    public override void Write(YamlWriter writer, sbyte value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlInt16Converter : YamlConverter<short> {
    public static YamlInt16Converter Instance { get; } = new YamlInt16Converter();
    public override short Read(YamlEventCursor cursor, YamlReadState state) => (short)YamlScalarText.ReadInteger(cursor, state, typeof(short), short.MinValue, short.MaxValue);
    public override void Write(YamlWriter writer, short value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlUInt16Converter : YamlConverter<ushort> {
    public static YamlUInt16Converter Instance { get; } = new YamlUInt16Converter();
    public override ushort Read(YamlEventCursor cursor, YamlReadState state) => (ushort)YamlScalarText.ReadInteger(cursor, state, typeof(ushort), ushort.MinValue, ushort.MaxValue);
    public override void Write(YamlWriter writer, ushort value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlInt32Converter : YamlConverter<int> {
    public static YamlInt32Converter Instance { get; } = new YamlInt32Converter();
    public override int Read(YamlEventCursor cursor, YamlReadState state) => (int)YamlScalarText.ReadInteger(cursor, state, typeof(int), int.MinValue, int.MaxValue);
    public override void Write(YamlWriter writer, int value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlUInt32Converter : YamlConverter<uint> {
    public static YamlUInt32Converter Instance { get; } = new YamlUInt32Converter();
    public override uint Read(YamlEventCursor cursor, YamlReadState state) => (uint)YamlScalarText.ReadInteger(cursor, state, typeof(uint), uint.MinValue, uint.MaxValue);
    public override void Write(YamlWriter writer, uint value, YamlWriteState state) => writer.WriteScalar((long)value);
}

internal sealed class YamlInt64Converter : YamlConverter<long> {
    public static YamlInt64Converter Instance { get; } = new YamlInt64Converter();
    public override long Read(YamlEventCursor cursor, YamlReadState state) => YamlScalarText.ReadInteger(cursor, state, typeof(long), long.MinValue, long.MaxValue);
    public override void Write(YamlWriter writer, long value, YamlWriteState state) => writer.WriteScalar(value);
}

internal sealed class YamlUInt64Converter : YamlConverter<ulong> {
    public static YamlUInt64Converter Instance { get; } = new YamlUInt64Converter();

    public override ulong Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(ulong));
        if (YamlSchemaResolver.TryParseInt64(text, cursor.Schema, out long value)) {
            if (value < 0) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(ulong), "outside the range of the type");
            return (ulong)value;
        }
        if (YamlBigIntegerConverter.TryParse(text, out BigInteger big) && big >= 0 && big <= ulong.MaxValue) return (ulong)big;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(ulong));
    }

    public override void Write(YamlWriter writer, ulong value, YamlWriteState state) {
        if (value <= long.MaxValue) writer.WriteScalar((long)value);
        else YamlScalarText.WriteInvariant(writer, value, null);
    }
}

internal sealed class YamlBigIntegerConverter : YamlConverter<BigInteger> {
    public static YamlBigIntegerConverter Instance { get; } = new YamlBigIntegerConverter();

    public override BigInteger Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(BigInteger));
        if (YamlSchemaResolver.TryParseInt64(text, cursor.Schema, out long value)) return value;
        if (TryParse(text, out BigInteger big)) return big;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(BigInteger));
    }

    public override void Write(YamlWriter writer, BigInteger value, YamlWriteState state) => YamlScalarText.WriteInvariant(writer, value, "D");

    /// <summary>Parses optionally signed decimal digits, the form integers too large for 64 bits are written in.</summary>
    public static bool TryParse(ReadOnlySpan<char> text, out BigInteger value) {
        value = default;
        int start = text.Length > 0 && (text[0] == '-' || text[0] == '+') ? 1 : 0;
        if (text.Length == start) return false;
        for (int i = start; i < text.Length; i++) {
            if (text[i] < '0' || text[i] > '9') return false;
        }
        return BigInteger.TryParse(text.ToString(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out value);
    }
}

internal sealed class YamlDoubleConverter : YamlConverter<double> {
    public static YamlDoubleConverter Instance { get; } = new YamlDoubleConverter();
    public override double Read(YamlEventCursor cursor, YamlReadState state) => YamlScalarText.ReadDouble(cursor, state, typeof(double));
    public override void Write(YamlWriter writer, double value, YamlWriteState state) => writer.WriteScalar(value);
}

internal sealed class YamlSingleConverter : YamlConverter<float> {
    public static YamlSingleConverter Instance { get; } = new YamlSingleConverter();

    public override float Read(YamlEventCursor cursor, YamlReadState state) {
        double value = YamlScalarText.ReadDouble(cursor, state, typeof(float));
        if (double.IsNaN(value) || double.IsInfinity(value)) return (float)value;
        // Parse the text as a float too, so values such as 0.1 round once to the nearest float, not through a double.
        string text = cursor.GetString();
        if (float.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out float exact)) return exact;
        return (float)value;
    }

    public override void Write(YamlWriter writer, float value, YamlWriteState state) => writer.WritePlainValue(YamlScalarFormat.Format(value).AsSpan());
}

internal sealed class YamlDecimalConverter : YamlConverter<decimal> {
    public static YamlDecimalConverter Instance { get; } = new YamlDecimalConverter();

    public override decimal Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(decimal));
        if (!YamlSchemaResolver.TryParseDecimal(text, cursor.Schema, out decimal value)) {
            throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(decimal), "not a number in the range of the type");
        }
        return value;
    }

    public override void Write(YamlWriter writer, decimal value, YamlWriteState state) => YamlScalarText.WriteInvariant(writer, value, null);
}

internal sealed class YamlCharConverter : YamlConverter<char> {
    public static YamlCharConverter Instance { get; } = new YamlCharConverter();

    public override char Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(char));
        if (text.Length != 1) throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(char), "not a single character");
        return text[0];
    }

    public override void Write(YamlWriter writer, char value, YamlWriteState state) => writer.WriteScalar(value.ToString());
}

/// <summary>
/// <see cref="DateTime"/> as an ISO 8601 timestamp, plain like a YAML timestamp: <c>2024-05-01T08:30:00Z</c>. Fractions
/// are written only when present, and the kind is kept: <c>Z</c> for UTC, an offset for local time, nothing for unspecified.
/// </summary>
internal sealed class YamlDateTimeConverter : YamlConverter<DateTime> {
    public const string Format = "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK";
    public static YamlDateTimeConverter Instance { get; } = new YamlDateTimeConverter();

    public override DateTime Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(DateTime)).ToString();
        if (DateTime.TryParse(Normalize(text), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(DateTime), "not an ISO 8601 timestamp");
    }

    public override void Write(YamlWriter writer, DateTime value, YamlWriteState state) => YamlScalarText.WriteInvariant(writer, value, Format);

    /// <summary>Accepts the YAML 1.1 timestamp spellings: a lower-case <c>t</c> or a space between date and time.</summary>
    public static string Normalize(string text) {
        if (text.Length > 10 && (text[10] == 't' || text[10] == ' ')) return text.Substring(0, 10) + "T" + text.Substring(11).Trim();
        return text;
    }
}

/// <summary><see cref="DateTimeOffset"/> as an ISO 8601 timestamp with its offset: <c>2024-05-01T08:30:00+02:00</c>.</summary>
internal sealed class YamlDateTimeOffsetConverter : YamlConverter<DateTimeOffset> {
    public const string Format = "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz";
    public static YamlDateTimeOffsetConverter Instance { get; } = new YamlDateTimeOffsetConverter();

    public override DateTimeOffset Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(DateTimeOffset)).ToString();
        if (DateTimeOffset.TryParse(YamlDateTimeConverter.Normalize(text), CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out DateTimeOffset value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(DateTimeOffset), "not an ISO 8601 timestamp");
    }

    public override void Write(YamlWriter writer, DateTimeOffset value, YamlWriteState state) => YamlScalarText.WriteInvariant(writer, value, Format);
}

/// <summary>
/// <see cref="TimeSpan"/> in the invariant constant format, <c>[-][d.]hh:mm:ss[.fffffff]</c>. It is written as a string,
/// so it is quoted where YAML 1.1 readers would take it for a sexagesimal number (<c>'12:30:00'</c>, but <c>08:30:00</c>).
/// </summary>
internal sealed class YamlTimeSpanConverter : YamlConverter<TimeSpan> {
    public static YamlTimeSpanConverter Instance { get; } = new YamlTimeSpanConverter();

    public override TimeSpan Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(TimeSpan)).ToString();
        if (TimeSpan.TryParseExact(text, "c", CultureInfo.InvariantCulture, out TimeSpan value)
            || TimeSpan.TryParse(text, CultureInfo.InvariantCulture, out value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TimeSpan), "not a time span such as 01:30:00");
    }

    public override void Write(YamlWriter writer, TimeSpan value, YamlWriteState state) => writer.WriteScalar(value.ToString("c", CultureInfo.InvariantCulture));
}

/// <summary><see cref="Guid"/> in the hyphenated <c>D</c> format.</summary>
internal sealed class YamlGuidConverter : YamlConverter<Guid> {
    public static YamlGuidConverter Instance { get; } = new YamlGuidConverter();

    public override Guid Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(Guid)).ToString();
        if (Guid.TryParse(text, out Guid value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(Guid));
    }

    public override void Write(YamlWriter writer, Guid value, YamlWriteState state) => writer.WriteScalar(value.ToString("D"));
}

/// <summary><see cref="Uri"/> as its original string; relative URIs are allowed.</summary>
internal sealed class YamlUriConverter : YamlConverter<Uri> {
    public static YamlUriConverter Instance { get; } = new YamlUriConverter();

    public override Uri Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(Uri)).ToString();
        if (Uri.TryCreate(text, UriKind.RelativeOrAbsolute, out Uri? value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(Uri));
    }

    public override void Write(YamlWriter writer, Uri value, YamlWriteState state) => writer.WriteScalar(value.OriginalString);
}

/// <summary>A byte array as a base64 scalar tagged <c>!!binary</c>, the YAML binary type. Reading accepts it with or without the tag.</summary>
internal sealed class YamlByteArrayConverter : YamlConverter<byte[]> {
    public const string BinaryTag = YamlSchemaResolver.TagPrefix + "binary";
    public static YamlByteArrayConverter Instance { get; } = new YamlByteArrayConverter();

    public override byte[] Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(byte[])).ToString();
        try {
            return Convert.FromBase64String(text);
        } catch (FormatException) {
            throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(byte[]), "not base64");
        }
    }

    public override void Write(YamlWriter writer, byte[] value, YamlWriteState state) {
        writer.WriteTag(BinaryTag);
        if (value.Length == 0) writer.WriteScalar(string.Empty);
        else writer.WritePlainValue(Convert.ToBase64String(value).AsSpan());
    }
}

#if !NET6_0_OR_GREATER
/// <summary>
/// <c>DateOnly</c> and <c>TimeOnly</c> for the .NET Standard and .NET Framework builds, which cannot name those types
/// but run on .NET 6 or later when an application targets an older framework than this package's .NET 8 build. The
/// formats match the .NET 8 converters; the parse method is found by reflection, which those builds do not trim.
/// </summary>
internal sealed class YamlFormattedValueConverter<T> : YamlConverter<T> where T : IFormattable {
    private readonly string _format;
    private readonly bool _plain;
    private readonly System.Reflection.MethodInfo _parse;

    public YamlFormattedValueConverter(string format, bool plain) {
        _format = format;
        _plain = plain;
        _parse = typeof(T).GetMethod("ParseExact", new[] { typeof(string), typeof(string), typeof(IFormatProvider), typeof(DateTimeStyles) })
            ?? throw new NotSupportedException("The type " + typeof(T).Name + " has no ParseExact method.");
    }

    /// <summary>Creates the converter's metadata for <c>DateOnly</c> or <c>TimeOnly</c>, or returns null for other types.</summary>
    public static YamlTypeInfo? TryCreate(Type type, YamlSerializerOptions options) {
        string? name = type.FullName;
        if (name != "System.DateOnly" && name != "System.TimeOnly") return null;
        bool date = name == "System.DateOnly";
        object converter = Activator.CreateInstance(typeof(YamlFormattedValueConverter<>).MakeGenericType(type), date ? "yyyy-MM-dd" : "HH':'mm':'ss.FFFFFFF", date)!;
        return (YamlTypeInfo)Activator.CreateInstance(typeof(YamlTypeInfo<>).MakeGenericType(type),
            System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic, null, new object?[] { options, converter, null }, null)!;
    }

    public override T Read(YamlEventCursor cursor, YamlReadState state) {
        string text = YamlScalarText.Read(cursor, state, typeof(T)).ToString();
        string[] formats = _plain ? new[] { _format } : new[] { _format, "HH':'mm':'ss", "HH':'mm" };
        foreach (string format in formats) {
            try {
                return (T)_parse.Invoke(null, new object[] { text, format, CultureInfo.InvariantCulture, DateTimeStyles.None })!;
            } catch (System.Reflection.TargetInvocationException exception) when (exception.InnerException is FormatException) {
            }
        }
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(T));
    }

    public override void Write(YamlWriter writer, T value, YamlWriteState state) {
        string text = value.ToString(_format, CultureInfo.InvariantCulture);
        if (_plain) writer.WritePlainValue(text.AsSpan());
        else writer.WriteScalar(text.TrimEnd('.'));
    }
}
#endif

#if NET6_0_OR_GREATER
/// <summary><see cref="DateOnly"/> as an ISO 8601 date, plain like a YAML timestamp: <c>2024-05-01</c>.</summary>
internal sealed class YamlDateOnlyConverter : YamlConverter<DateOnly> {
    public static YamlDateOnlyConverter Instance { get; } = new YamlDateOnlyConverter();

    public override DateOnly Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(DateOnly));
        if (DateOnly.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(DateOnly), "not a date such as 2024-05-01");
    }

    public override void Write(YamlWriter writer, DateOnly value, YamlWriteState state) {
        Span<char> buffer = stackalloc char[16];
        value.TryFormat(buffer, out int length, "yyyy-MM-dd", CultureInfo.InvariantCulture);
        writer.WritePlainValue(buffer.Slice(0, length));
    }
}

/// <summary><see cref="TimeOnly"/> as <c>HH:mm:ss[.fffffff]</c>, written as a string, so it is quoted where YAML 1.1 readers would take it for a number.</summary>
internal sealed class YamlTimeOnlyConverter : YamlConverter<TimeOnly> {
    public static YamlTimeOnlyConverter Instance { get; } = new YamlTimeOnlyConverter();

    public override TimeOnly Read(YamlEventCursor cursor, YamlReadState state) {
        ReadOnlySpan<char> text = YamlScalarText.Read(cursor, state, typeof(TimeOnly));
        if (TimeOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly value)) return value;
        throw YamlSerializationErrors.TypeMismatch(cursor, state, typeof(TimeOnly), "not a time such as 08:30:00");
    }

    public override void Write(YamlWriter writer, TimeOnly value, YamlWriteState state) =>
        writer.WriteScalar(value.ToString("HH':'mm':'ss.FFFFFFF", CultureInfo.InvariantCulture).TrimEnd('.'));
}
#endif

/// <summary>A nullable value type: null when the node is null, otherwise the underlying type's value.</summary>
internal sealed class YamlNullableConverter<T> : YamlConverter<T?> where T : struct {
    private YamlTypeInfo<T>? _underlying;

    public override void Configure(YamlSerializerOptions options) => _underlying = (YamlTypeInfo<T>)options.GetTypeInfo(typeof(T));

    public override T? Read(YamlEventCursor cursor, YamlReadState state) => _underlying!.Read(cursor, state);

    public override void Write(YamlWriter writer, T? value, YamlWriteState state) => _underlying!.Write(writer, value!.Value, state);
}
