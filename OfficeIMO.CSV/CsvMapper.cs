#nullable enable

using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Linq.Expressions;
using System.Reflection;
#if NET8_0_OR_GREATER
using System.Runtime.CompilerServices;
#endif

namespace OfficeIMO.CSV;

/// <summary>
/// Fluent mapping builder used to project CSV rows into typed models.
/// </summary>
public sealed class CsvMapper<T> where T : new()
{
    internal List<ICsvMappingEntry<T>> Entries { get; } = new();

    /// <summary>
    /// Binds a CSV column to an assignment delegate.
    /// </summary>
    public CsvMapper<T> FromColumn<TValue>(string columnName, Func<T, TValue, T> assign)
    {
        if (string.IsNullOrWhiteSpace(columnName))
        {
            throw new ArgumentException("Column name cannot be null or empty.", nameof(columnName));
        }

        if (assign is null)
        {
            throw new ArgumentNullException(nameof(assign));
        }

        Entries.Add(new CsvMappingEntry<T, TValue>(columnName, assign));
        return this;
    }
}

internal interface ICsvMappingEntry<T>
{
    string ColumnName { get; }

    T Apply(T instance, object? rawValue, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats);
}

internal sealed class CsvMappingEntry<T, TValue> : ICsvMappingEntry<T>
{
    public CsvMappingEntry(string columnName, Func<T, TValue, T> assign)
    {
        ColumnName = columnName;
        _assign = assign;
    }

    public string ColumnName { get; }

    public T Apply(T instance, object? rawValue, CultureInfo culture, IReadOnlyList<string>? dateTimeFormats)
    {
        var value = CsvValueConverter.ConvertTo<TValue>(rawValue, culture, dateTimeFormats);
        return _assign(instance, value!);
    }

    private readonly Func<T, TValue, T> _assign;
}

/// <summary>
/// Extension methods enabling typed mapping projections.
/// </summary>
public static class CsvMappingExtensions
{
    /// <summary>
    /// Projects the remaining unread rows of a forward-only data reader into
    /// instances of <typeparamref name="T"/> by matching column names to writable
    /// public properties. The caller retains ownership of the reader.
    /// </summary>
    public static IEnumerable<T> RowsAs<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this DbDataReader reader) where T : new()
    {
        if (reader is null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        return EnumerateAutomaticReaderRows<T>(reader);
    }

    /// <summary>
    /// Projects the remaining unread rows of a forward-only data reader into
    /// instances of <typeparamref name="T"/> using explicit, AOT-friendly column
    /// assignments. The caller retains ownership of the reader.
    /// </summary>
    public static IEnumerable<T> RowsAs<T>(
        this DbDataReader reader,
        Action<CsvMapper<T>> configure) where T : new()
    {
        if (reader is null)
        {
            throw new ArgumentNullException(nameof(reader));
        }

        return EnumerateMappedReaderRows(reader, configure);
    }

    /// <summary>
    /// Projects rows into instances of <typeparamref name="T"/> by matching CSV
    /// headers to writable public properties. Matching is case-insensitive and
    /// also ignores spaces and punctuation.
    /// </summary>
    public static IEnumerable<T> RowsAs<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this CsvDocument document) where T : new()
    {
        if (document is null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        return EnumerateAutomaticRows<T>(document);
    }

    /// <summary>
    /// Projects rows into instances of <typeparamref name="T"/> using explicit,
    /// AOT-friendly column assignments.
    /// </summary>
    public static IEnumerable<T> RowsAs<T>(
        this CsvDocument document,
        Action<CsvMapper<T>> configure) where T : new() => EnumerateMappedRows(document, configure);

    private static IEnumerable<T> EnumerateMappedRows<T>(CsvDocument document, Action<CsvMapper<T>> configure) where T : new()
    {
        if (document is null)
        {
            throw new ArgumentNullException(nameof(document));
        }

        if (configure is null)
        {
            throw new ArgumentNullException(nameof(configure));
        }

        var mapper = new CsvMapper<T>();
        configure(mapper);
        if (mapper.Entries.Count == 0)
        {
            yield break;
        }

        var bindings = mapper.Entries
            .Select(entry => new MappingBinding<T>(document.GetColumnIndex(entry.ColumnName), entry))
            .ToArray();

        foreach (var row in document.AsEnumerable())
        {
            var instance = new T();
            foreach (var binding in bindings)
            {
                var rawValue = row[binding.ColumnIndex];
                instance = binding.Entry.Apply(instance, rawValue, document.Culture, document.DateTimeFormats);
            }

            yield return instance;
        }
    }

    private readonly record struct MappingBinding<T>(int ColumnIndex, ICsvMappingEntry<T> Entry);

    private static IEnumerable<T> EnumerateMappedReaderRows<T>(
        DbDataReader reader,
        Action<CsvMapper<T>> configure) where T : new()
    {
        if (configure is null)
        {
            throw new ArgumentNullException(nameof(configure));
        }

        var mapper = new CsvMapper<T>();
        configure(mapper);
        if (mapper.Entries.Count == 0)
        {
            yield break;
        }

        var bindings = mapper.Entries
            .Select(entry => new MappingBinding<T>(reader.GetOrdinal(entry.ColumnName), entry))
            .ToArray();
        GetReaderConversionOptions(reader, out CultureInfo culture, out IReadOnlyList<string>? dateTimeFormats);

        while (reader.Read())
        {
            T instance = new T();
            foreach (MappingBinding<T> binding in bindings)
            {
                object? rawValue = reader.GetValue(binding.ColumnIndex);
                instance = binding.Entry.Apply(
                    instance,
                    ReferenceEquals(rawValue, DBNull.Value) ? null : rawValue,
                    culture,
                    dateTimeFormats);
            }

            yield return instance;
        }
    }

    private static AutomaticMappingBinding<T>[] CreateAutomaticBindings<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IReadOnlyList<string> headers)
        where T : new()
    {
        var mapped = new AutomaticPropertySetter<T>?[headers.Count];
        var assigned = new HashSet<AutomaticPropertySetter<T>>();
        for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
        {
            string header = headers[columnIndex];
            if (TryResolveUnassignedProperty(
                    AutomaticMappingCache<T>.Exact,
                    header,
                    assigned,
                    header,
                    "with that exact name",
                    out AutomaticPropertySetter<T> property))
            {
                mapped[columnIndex] = property;
                assigned.Add(property);
            }
        }

        for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
        {
            if (mapped[columnIndex] is not null)
            {
                continue;
            }

            string header = headers[columnIndex];
            if (TryResolveUnassignedProperty(
                    AutomaticMappingCache<T>.Insensitive,
                    header,
                    assigned,
                    header,
                    "when casing is ignored",
                    out AutomaticPropertySetter<T> property))
            {
                mapped[columnIndex] = property;
                assigned.Add(property);
            }
        }

        for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
        {
            if (mapped[columnIndex] is not null)
            {
                continue;
            }

            string key = Canonicalize(headers[columnIndex]);
            if (key.Length == 0)
            {
                continue;
            }

            if (TryResolveUnassignedProperty(
                    AutomaticMappingCache<T>.Canonical,
                    key,
                    assigned,
                    headers[columnIndex],
                    "when punctuation is ignored",
                    out AutomaticPropertySetter<T> property))
            {
                mapped[columnIndex] = property;
                assigned.Add(property);
            }
        }

        var bindings = new List<AutomaticMappingBinding<T>>(mapped.Length);
        for (var columnIndex = 0; columnIndex < mapped.Length; columnIndex++)
        {
            if (mapped[columnIndex] is { } property)
            {
                bindings.Add(new AutomaticMappingBinding<T>(columnIndex, property));
            }
        }

        return bindings.ToArray();
    }

    private static bool TryResolveUnassignedProperty<T>(
        IReadOnlyDictionary<string, AutomaticPropertySetter<T>[]> lookup,
        string key,
        HashSet<AutomaticPropertySetter<T>> assigned,
        string header,
        string ambiguityDescription,
        out AutomaticPropertySetter<T> property)
    {
        property = null!;
        if (!lookup.TryGetValue(key, out AutomaticPropertySetter<T>[]? candidates))
        {
            return false;
        }

        for (var index = 0; index < candidates.Length; index++)
        {
            AutomaticPropertySetter<T> candidate = candidates[index];
            if (assigned.Contains(candidate))
            {
                continue;
            }

            if (property is not null)
            {
                throw new CsvException(
                    $"CSV header '{header}' matches multiple writable properties on {typeof(T).Name} {ambiguityDescription}.");
            }

            property = candidate;
        }

        return property is not null;
    }

    private static IEnumerable<T> EnumerateAutomaticRows<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        CsvDocument document) where T : new()
    {
        AutomaticMappingBinding<T>[] bindings = CreateAutomaticBindings<T>(document.Header);
        if (bindings.Length == 0)
        {
            throw new CsvException($"No CSV headers match writable properties on {typeof(T).Name}.");
        }

        foreach (CsvRow row in document.AsEnumerable())
        {
            T instance = new T();
            for (var index = 0; index < bindings.Length; index++)
            {
                AutomaticMappingBinding<T> binding = bindings[index];
                object? rawValue = row[binding.ColumnIndex];
                if (!CsvValueConverter.TryConvert(
                        rawValue,
                        binding.Property.ValueType,
                        document.Culture,
                        document.DateTimeFormats,
                        out object? converted,
                        out string? error))
                {
                    throw new CsvException(
                        $"Column '{document.Header[binding.ColumnIndex]}' cannot be assigned to " +
                        $"{typeof(T).Name}.{binding.Property.Name}: {error}");
                }

                instance = binding.Property.Assign(instance, converted);
            }

            yield return instance;
        }
    }

    private static IEnumerable<T> EnumerateAutomaticReaderRows<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        DbDataReader reader) where T : new()
    {
        var headers = new string[reader.FieldCount];
        for (var index = 0; index < headers.Length; index++)
        {
            headers[index] = reader.GetName(index);
        }

        AutomaticMappingBinding<T>[] bindings = CreateAutomaticBindings<T>(headers);
        if (bindings.Length == 0)
        {
            throw new CsvException($"No reader columns match writable properties on {typeof(T).Name}.");
        }

        GetReaderConversionOptions(reader, out CultureInfo culture, out IReadOnlyList<string>? dateTimeFormats);

        while (reader.Read())
        {
            T instance = new T();
            for (var index = 0; index < bindings.Length; index++)
            {
                AutomaticMappingBinding<T> binding = bindings[index];
                object? rawValue = reader.GetValue(binding.ColumnIndex);
                if (ReferenceEquals(rawValue, DBNull.Value))
                {
                    rawValue = null;
                }

                if (!CsvValueConverter.TryConvert(
                        rawValue,
                        binding.Property.ValueType,
                        culture,
                        dateTimeFormats,
                        out object? converted,
                        out string? error))
                {
                    throw new CsvException(
                        $"Column '{headers[binding.ColumnIndex]}' cannot be assigned to " +
                        $"{typeof(T).Name}.{binding.Property.Name}: {error}");
                }

                instance = binding.Property.Assign(instance, converted);
            }

            yield return instance;
        }
    }

    private static void GetReaderConversionOptions(
        DbDataReader reader,
        out CultureInfo culture,
        out IReadOnlyList<string>? dateTimeFormats)
    {
        if (reader is CsvDataReader csvReader)
        {
            culture = csvReader.MappingCulture;
            dateTimeFormats = csvReader.MappingDateTimeFormats;
            return;
        }

        culture = CultureInfo.InvariantCulture;
        dateTimeFormats = null;
    }

    private static class AutomaticMappingCache<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T> where T : new()
    {
        internal static readonly Dictionary<string, AutomaticPropertySetter<T>[]> Exact;
        internal static readonly Dictionary<string, AutomaticPropertySetter<T>[]> Insensitive;
        internal static readonly Dictionary<string, AutomaticPropertySetter<T>[]> Canonical;

        static AutomaticMappingCache()
        {
            AutomaticPropertySetter<T>[] properties = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(static property => property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
                .Select(static property => new AutomaticPropertySetter<T>(property, CreateAssignment<T>(property)))
                .ToArray();

            Exact = CreateLookup(
                properties,
                static property => property.Name,
                StringComparer.Ordinal);
            Insensitive = CreateLookup(
                properties,
                static property => property.Name,
                StringComparer.OrdinalIgnoreCase);
            Canonical = CreateLookup(
                properties,
                static property => Canonicalize(property.Name),
                StringComparer.Ordinal);
        }
    }

    private static Dictionary<string, AutomaticPropertySetter<T>[]> CreateLookup<T>(
        IEnumerable<AutomaticPropertySetter<T>> properties,
        Func<AutomaticPropertySetter<T>, string> getKey,
        IEqualityComparer<string> comparer)
    {
        return properties
            .Select(property => new { Key = getKey(property), Property = property })
            .Where(static item => item.Key.Length > 0)
            .GroupBy(static item => item.Key, comparer)
            .ToDictionary(
                static group => group.Key,
                static group => group.Select(static item => item.Property).ToArray(),
                comparer);
    }

    private sealed class AutomaticPropertySetter<T>
    {
        internal AutomaticPropertySetter(PropertyInfo property, Func<T, object?, T> assign)
        {
            Name = property.Name;
            ValueType = property.PropertyType;
            Assign = assign;
        }

        internal string Name { get; }

        internal Type ValueType { get; }

        internal Func<T, object?, T> Assign { get; }
    }

    private static Func<T, object?, T> CreateAssignment<T>(PropertyInfo property)
    {
#if NET8_0_OR_GREATER
        if (!RuntimeFeature.IsDynamicCodeSupported)
        {
            return CreateReflectionAssignment<T>(property);
        }
#endif

        try
        {
            ParameterExpression target = Expression.Parameter(typeof(T), "target");
            ParameterExpression value = Expression.Parameter(typeof(object), "value");
            BinaryExpression assignment = Expression.Assign(
                Expression.Property(target, property),
                Expression.Convert(value, property.PropertyType));
            BlockExpression body = Expression.Block(assignment, target);
            return Expression.Lambda<Func<T, object?, T>>(body, target, value).Compile();
        }
        catch
        {
            return CreateReflectionAssignment<T>(property);
        }
    }

    private static Func<T, object?, T> CreateReflectionAssignment<T>(PropertyInfo property)
    {
        return (target, value) =>
        {
            object boxed = target!;
            property.SetValue(boxed, value, index: null);
            return (T)boxed;
        };
    }

    private static string Canonicalize(string value)
    {
        var builder = new System.Text.StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            char character = value[index];
            if (char.IsLetterOrDigit(character))
            {
                builder.Append(char.ToUpperInvariant(character));
            }
        }

        return builder.ToString();
    }

    private readonly record struct AutomaticMappingBinding<T>(int ColumnIndex, AutomaticPropertySetter<T> Property);
}
