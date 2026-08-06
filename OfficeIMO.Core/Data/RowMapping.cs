#nullable enable

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;

namespace OfficeIMO.Data;

/// <summary>Supplies conversion metadata for a tabular data reader.</summary>
public interface IDataReaderMappingMetadata {
    /// <summary>Gets the culture used when text values are converted to typed properties.</summary>
    CultureInfo MappingCulture { get; }

    /// <summary>Gets optional exact date and time formats used during typed conversion.</summary>
    IReadOnlyList<string>? MappingDateTimeFormats { get; }

    /// <summary>Gets an optional converter invoked before the built-in typed conversion pipeline.</summary>
    Func<object, Type, CultureInfo, (bool ok, object? value)>? MappingTypeConverter { get; }

    /// <summary>Gets whether every non-empty source column must resolve to a writable property.</summary>
    bool RequireAllColumnsMapped { get; }
}

/// <summary>Supplies additional column aliases for a model property.</summary>
public interface IDataColumnAliasProvider {
    /// <summary>Gets the column names that may bind to the decorated property.</summary>
    IReadOnlyList<string> ColumnAliases { get; }
}

/// <summary>Reports a failure while mapping tabular values to a typed model.</summary>
public sealed class DataMappingException : InvalidOperationException {
    /// <summary>Initializes a mapping exception with a descriptive message.</summary>
    public DataMappingException(string message) : base(message) { }
}

/// <summary>Defines explicit, AOT-friendly column assignments for a typed row.</summary>
public sealed class RowMapper<T> where T : new() {
    internal List<IRowMappingEntry<T>> Entries { get; } = new();

    /// <summary>Binds a named column to an assignment delegate.</summary>
    public RowMapper<T> FromColumn<TValue>(string columnName, Func<T, TValue, T> assign) {
        if (string.IsNullOrWhiteSpace(columnName)) {
            throw new ArgumentException("Column name cannot be null or empty.", nameof(columnName));
        }
        if (assign is null) {
            throw new ArgumentNullException(nameof(assign));
        }

        Entries.Add(new RowMappingEntry<T, TValue>(columnName, assign));
        return this;
    }
}

/// <summary>Typed row projections for any forward-only <see cref="DbDataReader"/>.</summary>
public static class DataReaderMappingExtensions {
    /// <summary>
    /// Projects the remaining unread rows by matching column names to writable public properties.
    /// The caller retains ownership of the reader.
    /// </summary>
    public static IEnumerable<T> RowsAs<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this DbDataReader reader) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        return EnumerateAutomatic<T>(reader);
    }

    /// <summary>
    /// Projects the remaining unread rows using explicit column assignments.
    /// The caller retains ownership of the reader.
    /// </summary>
    public static IEnumerable<T> RowsAs<T>(
        this DbDataReader reader,
        Action<RowMapper<T>> configure) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        return EnumerateExplicit(reader, configure);
    }

    private static IEnumerable<T> EnumerateAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        DbDataReader reader) where T : new() {
        if (reader.FieldCount == 0) yield break;

        GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out bool requireAllColumnsMapped);
        AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(GetHeaders(reader), requireAllColumnsMapped);
        while (reader.Read()) {
            yield return plan.MapRow(
                index => NormalizeDatabaseNull(reader.GetValue(index)),
                culture,
                dateTimeFormats,
                typeConverter);
        }
    }

    private static IEnumerable<T> EnumerateExplicit<T>(
        DbDataReader reader,
        Action<RowMapper<T>> configure) where T : new() {
        if (reader.FieldCount == 0) yield break;

        ExplicitRowMappingPlan<T> plan = ExplicitRowMappingPlan<T>.Create(GetHeaders(reader), configure);
        if (plan.IsEmpty) yield break;

        GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out _);
        while (reader.Read()) {
            yield return plan.MapRow(
                index => NormalizeDatabaseNull(reader.GetValue(index)),
                culture,
                dateTimeFormats,
                typeConverter);
        }
    }

    private static string[] GetHeaders(DbDataReader reader) {
        var headers = new string[reader.FieldCount];
        for (int index = 0; index < headers.Length; index++) {
            headers[index] = reader.GetName(index);
        }
        return headers;
    }

    private static object? NormalizeDatabaseNull(object? value) =>
        ReferenceEquals(value, DBNull.Value) ? null : value;

    private static void GetConversionOptions(
        DbDataReader reader,
        out CultureInfo culture,
        out IReadOnlyList<string>? dateTimeFormats,
        out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
        out bool requireAllColumnsMapped) {
        if (reader is IDataReaderMappingMetadata metadata) {
            culture = metadata.MappingCulture;
            dateTimeFormats = metadata.MappingDateTimeFormats;
            typeConverter = metadata.MappingTypeConverter;
            requireAllColumnsMapped = metadata.RequireAllColumnsMapped;
            return;
        }

        culture = CultureInfo.InvariantCulture;
        dateTimeFormats = null;
        typeConverter = null;
        requireAllColumnsMapped = false;
    }
}

internal interface IRowMappingEntry<T> {
    string ColumnName { get; }
    T Apply(
        T instance,
        object? rawValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter);
}

internal sealed class RowMappingEntry<T, TValue> : IRowMappingEntry<T> {
    private readonly Func<T, TValue, T> _assign;

    internal RowMappingEntry(string columnName, Func<T, TValue, T> assign) {
        ColumnName = columnName;
        _assign = assign;
    }

    public string ColumnName { get; }

    public T Apply(
        T instance,
        object? rawValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter) {
        TValue? value = DataValueConverter.ConvertTo<TValue>(rawValue, culture, dateTimeFormats, typeConverter);
        return _assign(instance, value!);
    }
}

internal static class DataValueConverter {
    internal static T? ConvertTo<T>(
        object? value,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats = null,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null) {
        if (!TryConvert(value, typeof(T), culture, dateTimeFormats, typeConverter, out object? result, out string? error)) {
            throw new DataMappingException(error ?? $"Value '{value}' cannot be converted to {typeof(T).Name}.");
        }
        return (T?)result;
    }

    internal static bool TryConvert(
        object? value,
        Type targetType,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
        out object? result,
        out string? error) {
        error = null;
        result = null;
        Type? underlyingType = Nullable.GetUnderlyingType(targetType);
        Type effectiveType = underlyingType ?? targetType;

        if (value is null) {
            if (underlyingType is null && targetType.IsValueType) {
                error = $"Cannot assign null to non-nullable type {targetType.Name}.";
                return false;
            }
            return true;
        }
        if (typeConverter is not null) {
            try {
                (bool handled, object? converted) = typeConverter(value, effectiveType, culture);
                if (handled) {
                    result = converted;
                    return true;
                }
            } catch (Exception ex) {
                error = ex.Message;
                return false;
            }
        }
        if (effectiveType.IsInstanceOfType(value)) {
            result = value;
            return true;
        }
        if (value is string text) {
            return TryConvertFromString(text, effectiveType, culture, dateTimeFormats, out result, out error);
        }
        try {
            result = Convert.ChangeType(value, effectiveType, culture);
            return true;
        } catch (Exception ex) {
            error = ex.Message;
            return false;
        }
    }

    private static bool TryConvertFromString(
        string text,
        Type targetType,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        out object? result,
        out string? error) {
        error = null;
        result = null;
        try {
            if (targetType == typeof(string)) result = text;
            else if (targetType == typeof(int) && int.TryParse(text, NumberStyles.Any, culture, out int intValue)) result = intValue;
            else if (targetType == typeof(long) && long.TryParse(text, NumberStyles.Any, culture, out long longValue)) result = longValue;
            else if (targetType == typeof(short) && short.TryParse(text, NumberStyles.Any, culture, out short shortValue)) result = shortValue;
            else if (targetType == typeof(byte) && byte.TryParse(text, NumberStyles.Any, culture, out byte byteValue)) result = byteValue;
            else if (targetType == typeof(bool) && bool.TryParse(text, out bool boolValue)) result = boolValue;
            else if (targetType == typeof(bool) && text == "0") result = false;
            else if (targetType == typeof(bool) && text == "1") result = true;
            else if (targetType == typeof(double) && double.TryParse(text, NumberStyles.Any, culture, out double doubleValue)) result = doubleValue;
            else if (targetType == typeof(decimal) && decimal.TryParse(text, NumberStyles.Any, culture, out decimal decimalValue)) result = decimalValue;
            else if (targetType == typeof(float) && float.TryParse(text, NumberStyles.Any, culture, out float floatValue)) result = floatValue;
            else if (targetType == typeof(DateTime) && dateTimeFormats is { Count: > 0 } &&
                     DateTime.TryParseExact(text, dateTimeFormats as string[] ?? dateTimeFormats.ToArray(), culture, DateTimeStyles.None, out DateTime formattedDateTime)) result = formattedDateTime;
            else if (targetType == typeof(DateTime)) result = DateTime.Parse(text, culture, DateTimeStyles.None);
            else if (targetType == typeof(Guid) && Guid.TryParse(text, out Guid guidValue)) result = guidValue;
            else if (targetType.IsEnum) result = Enum.Parse(targetType, text, ignoreCase: true);
            else result = Convert.ChangeType(text, targetType, culture);
            return true;
        } catch (Exception ex) {
            error = ex.Message;
            return false;
        }
    }
}
