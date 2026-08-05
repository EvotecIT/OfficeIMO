#nullable enable

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.Serialization;
#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
#endif

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
        RowMappingPlan<T> plan = RowMappingPlan<T>.CreateAutomatic(GetHeaders(reader), requireAllColumnsMapped);
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

        RowMappingPlan<T> plan = RowMappingPlan<T>.CreateExplicit(GetHeaders(reader), configure);
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

internal sealed class RowMappingPlan<
    [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T> where T : new() {
    private readonly MappingBinding<T>[] _bindings;

    private RowMappingPlan(MappingBinding<T>[] bindings) {
        _bindings = bindings;
    }

    internal bool IsEmpty => _bindings.Length == 0;

    internal static RowMappingPlan<T> CreateAutomatic(IReadOnlyList<string> headers, bool requireAllColumnsMapped = false) {
        AutomaticPropertySetter<T>?[] mapped = new AutomaticPropertySetter<T>?[headers.Count];
        var assigned = new HashSet<AutomaticPropertySetter<T>>();
        MapPass(headers, mapped, assigned, AutomaticMappingCache<T>.Exact, static value => value, "with that exact name");
        MapPass(headers, mapped, assigned, AutomaticMappingCache<T>.Insensitive, static value => value, "when casing is ignored");
        MapPass(headers, mapped, assigned, AutomaticMappingCache<T>.Aliases, static value => value, "through a declared alias");
        MapPass(headers, mapped, assigned, AutomaticMappingCache<T>.Canonical, Canonicalize, "when punctuation is ignored");
        MapPass(headers, mapped, assigned, AutomaticMappingCache<T>.CanonicalAliases, Canonicalize, "through a normalized alias");

        if (requireAllColumnsMapped) {
            string[] unmappedHeaders = headers
                .Where((header, index) => mapped[index] is null && !string.IsNullOrWhiteSpace(header))
                .ToArray();
            if (unmappedHeaders.Length > 0) {
                throw new DataMappingException(
                    $"Typed mapping for '{typeof(T).Name}' is strict and could not resolve columns: {string.Join(", ", unmappedHeaders.Select(static header => $"'{header}'"))}.");
            }
        }

        MappingBinding<T>[] bindings = mapped
            .Select((property, index) => property is null
                ? default(MappingBinding<T>?)
                : new MappingBinding<T>(index, property))
            .Where(static binding => binding.HasValue)
            .Select(static binding => binding!.Value)
            .ToArray();

        if (bindings.Length == 0) {
            throw new DataMappingException($"No columns match writable properties on {typeof(T).Name}.");
        }
        return new RowMappingPlan<T>(bindings);
    }

    internal static RowMappingPlan<T> CreateExplicit(
        IReadOnlyList<string> headers,
        Action<RowMapper<T>> configure) {
        var mapper = new RowMapper<T>();
        configure(mapper);
        MappingBinding<T>[] bindings = mapper.Entries
            .Select(entry => new MappingBinding<T>(FindColumn(headers, entry.ColumnName), entry))
            .ToArray();
        return new RowMappingPlan<T>(bindings);
    }

    internal T MapRow(
        Func<int, object?> getValue,
        CultureInfo culture,
        IReadOnlyList<string>? dateTimeFormats,
        Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter = null) {
        T instance = new T();
        foreach (MappingBinding<T> binding in _bindings) {
            instance = binding.Apply(instance, getValue(binding.ColumnIndex), culture, dateTimeFormats, typeConverter);
        }
        return instance;
    }

    private static void MapPass(
        IReadOnlyList<string> headers,
        AutomaticPropertySetter<T>?[] mapped,
        HashSet<AutomaticPropertySetter<T>> assigned,
        IReadOnlyDictionary<string, AutomaticPropertySetter<T>[]> lookup,
        Func<string, string> keySelector,
        string ambiguityDescription) {
        for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++) {
            if (mapped[columnIndex] is not null) continue;
            string key = keySelector(headers[columnIndex]);
            if (key.Length == 0 || !lookup.TryGetValue(key, out AutomaticPropertySetter<T>[]? candidates)) continue;

            AutomaticPropertySetter<T>? match = null;
            foreach (AutomaticPropertySetter<T> candidate in candidates) {
                if (assigned.Contains(candidate)) continue;
                if (match is not null) {
                    throw new DataMappingException(
                        $"Column '{headers[columnIndex]}' matches multiple writable properties on {typeof(T).Name} {ambiguityDescription}.");
                }
                match = candidate;
            }
            if (match is null) continue;
            mapped[columnIndex] = match;
            assigned.Add(match);
        }
    }

    private static int FindColumn(IReadOnlyList<string> headers, string columnName) {
        int found = -1;
        for (int index = 0; index < headers.Count; index++) {
            if (!string.Equals(headers[index], columnName, StringComparison.OrdinalIgnoreCase)) continue;
            if (found >= 0) {
                throw new DataMappingException($"Column name '{columnName}' is ambiguous.");
            }
            found = index;
        }
        if (found < 0) throw new DataMappingException($"Column '{columnName}' was not found.");
        return found;
    }

    private static string Canonicalize(string value) {
        var builder = new System.Text.StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsLetterOrDigit(character)) builder.Append(char.ToUpperInvariant(character));
        }
        return builder.ToString();
    }

    private readonly struct MappingBinding<TTarget> where TTarget : new() {
        private readonly AutomaticPropertySetter<TTarget>? _property;
        private readonly IRowMappingEntry<TTarget>? _entry;

        internal MappingBinding(int columnIndex, AutomaticPropertySetter<TTarget> property) {
            ColumnIndex = columnIndex;
            _property = property;
            _entry = null;
        }

        internal MappingBinding(int columnIndex, IRowMappingEntry<TTarget> entry) {
            ColumnIndex = columnIndex;
            _property = null;
            _entry = entry;
        }

        internal int ColumnIndex { get; }

        internal TTarget Apply(
            TTarget instance,
            object? rawValue,
            CultureInfo culture,
            IReadOnlyList<string>? dateTimeFormats,
            Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter) {
            if (_entry is not null) return _entry.Apply(instance, rawValue, culture, dateTimeFormats, typeConverter);
            if (!DataValueConverter.TryConvert(rawValue, _property!.ValueType, culture, dateTimeFormats, typeConverter, out object? converted, out string? error)) {
                throw new DataMappingException(
                    $"Column value cannot be assigned to {typeof(TTarget).Name}.{_property.Name}: {error}");
            }
            return _property.Assign(instance, converted);
        }
    }

    private sealed class AutomaticPropertySetter<TTarget> {
        internal AutomaticPropertySetter(PropertyInfo property, Func<TTarget, object?, TTarget> assign) {
            Property = property;
            Name = property.Name;
            ValueType = property.PropertyType;
            Assign = assign;
        }

        internal PropertyInfo Property { get; }
        internal string Name { get; }
        internal Type ValueType { get; }
        internal Func<TTarget, object?, TTarget> Assign { get; }
    }

    private static class AutomaticMappingCache<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] TTarget> {
        internal static readonly Dictionary<string, AutomaticPropertySetter<TTarget>[]> Exact;
        internal static readonly Dictionary<string, AutomaticPropertySetter<TTarget>[]> Insensitive;
        internal static readonly Dictionary<string, AutomaticPropertySetter<TTarget>[]> Aliases;
        internal static readonly Dictionary<string, AutomaticPropertySetter<TTarget>[]> Canonical;
        internal static readonly Dictionary<string, AutomaticPropertySetter<TTarget>[]> CanonicalAliases;

        static AutomaticMappingCache() {
            AutomaticPropertySetter<TTarget>[] properties = typeof(TTarget)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(static property => property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                .Select(static property => new AutomaticPropertySetter<TTarget>(property, CreateAssignment<TTarget>(property)))
                .ToArray();
            Exact = CreateLookup(properties, static property => property.Name, StringComparer.Ordinal);
            Insensitive = CreateLookup(properties, static property => property.Name, StringComparer.OrdinalIgnoreCase);
            Aliases = CreateAliasLookup(properties, canonical: false);
            Canonical = CreateLookup(properties, static property => Canonicalize(property.Name), StringComparer.Ordinal);
            CanonicalAliases = CreateAliasLookup(properties, canonical: true);
        }
    }

    private static Dictionary<string, AutomaticPropertySetter<TTarget>[]> CreateAliasLookup<TTarget>(
        IEnumerable<AutomaticPropertySetter<TTarget>> properties,
        bool canonical) {
        IEqualityComparer<string> comparer = canonical ? StringComparer.Ordinal : StringComparer.OrdinalIgnoreCase;
        return properties
            .SelectMany(property => GetAliases(property.Property)
                .Select(alias => new { Key = canonical ? Canonicalize(alias) : alias, Property = property }))
            .Where(static item => item.Key.Length > 0)
            .GroupBy(static item => item.Key, comparer)
            .ToDictionary(
                static group => group.Key,
                static group => group.Select(static item => item.Property).Distinct().ToArray(),
                comparer);
    }

    private static IEnumerable<string> GetAliases(PropertyInfo property) {
        var aliases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (property.GetCustomAttribute<DisplayNameAttribute>(inherit: true)?.DisplayName is { } displayName &&
            !string.IsNullOrWhiteSpace(displayName)) {
            aliases.Add(displayName);
        }
        if (property.GetCustomAttribute<DataMemberAttribute>(inherit: true)?.Name is { } dataMemberName &&
            !string.IsNullOrWhiteSpace(dataMemberName)) {
            aliases.Add(dataMemberName);
        }
        foreach (IDataColumnAliasProvider provider in property
                     .GetCustomAttributes(inherit: true)
                     .OfType<IDataColumnAliasProvider>()) {
            foreach (string alias in provider.ColumnAliases) {
                if (!string.IsNullOrWhiteSpace(alias)) aliases.Add(alias);
            }
        }
        return aliases;
    }

    private static Dictionary<string, AutomaticPropertySetter<TTarget>[]> CreateLookup<TTarget>(
        IEnumerable<AutomaticPropertySetter<TTarget>> properties,
        Func<AutomaticPropertySetter<TTarget>, string> getKey,
        IEqualityComparer<string> comparer) {
        return properties
            .Select(property => new { Key = getKey(property), Property = property })
            .Where(static item => item.Key.Length > 0)
            .GroupBy(static item => item.Key, comparer)
            .ToDictionary(
                static group => group.Key,
                static group => group.Select(static item => item.Property).ToArray(),
                comparer);
    }

    private static Func<TTarget, object?, TTarget> CreateAssignment<TTarget>(PropertyInfo property) {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported) return CreateReflectionAssignment<TTarget>(property);
#endif
        try {
            ParameterExpression target = Expression.Parameter(typeof(TTarget), "target");
            ParameterExpression value = Expression.Parameter(typeof(object), "value");
            BinaryExpression assignment = Expression.Assign(
                Expression.Property(target, property),
                Expression.Convert(value, property.PropertyType));
            BlockExpression body = Expression.Block(assignment, target);
            return Expression.Lambda<Func<TTarget, object?, TTarget>>(body, target, value).Compile();
        } catch {
            return CreateReflectionAssignment<TTarget>(property);
        }
    }

    private static Func<TTarget, object?, TTarget> CreateReflectionAssignment<TTarget>(PropertyInfo property) =>
        (target, value) => {
            object boxed = target!;
            property.SetValue(boxed, value, index: null);
            return (TTarget)boxed;
        };
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
