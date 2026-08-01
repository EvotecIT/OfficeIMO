#nullable enable

using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;

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
        Action<CsvMapper<T>> configure) where T : new() => document.Map(configure);

    /// <summary>
    /// Projects the document rows into a sequence of <typeparamref name="T"/> using the specified mapping configuration.
    /// </summary>
    public static IEnumerable<T> Map<T>(this CsvDocument document, Action<CsvMapper<T>> configure) where T : new()
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

    private static AutomaticMappingBinding[] CreateAutomaticBindings<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        IReadOnlyList<string> headers)
    {
        var writableProperties = typeof(T)
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(static property => property.CanWrite && property.GetIndexParameters().Length == 0)
            .ToArray();
        var exact = writableProperties.ToDictionary(static property => property.Name, StringComparer.OrdinalIgnoreCase);
        var canonical = new Dictionary<string, PropertyInfo>(StringComparer.Ordinal);
        foreach (var property in writableProperties)
        {
            string key = Canonicalize(property.Name);
            if (key.Length > 0 && !canonical.ContainsKey(key))
            {
                canonical.Add(key, property);
            }
        }

        var assigned = new HashSet<PropertyInfo>();
        var bindings = new List<AutomaticMappingBinding>();
        for (var columnIndex = 0; columnIndex < headers.Count; columnIndex++)
        {
            string header = headers[columnIndex];
            if (!exact.TryGetValue(header, out PropertyInfo? property))
            {
                canonical.TryGetValue(Canonicalize(header), out property);
            }

            if (property is not null && assigned.Add(property))
            {
                bindings.Add(new AutomaticMappingBinding(columnIndex, property));
            }
        }

        return bindings.ToArray();
    }

    private static IEnumerable<T> EnumerateAutomaticRows<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        CsvDocument document) where T : new()
    {
        AutomaticMappingBinding[] bindings = CreateAutomaticBindings<T>(document.Header);
        if (bindings.Length == 0)
        {
            throw new CsvException($"No CSV headers match writable properties on {typeof(T).Name}.");
        }

        foreach (CsvRow row in document.AsEnumerable())
        {
            object instance = new T()!;
            for (var index = 0; index < bindings.Length; index++)
            {
                AutomaticMappingBinding binding = bindings[index];
                object? rawValue = row[binding.ColumnIndex];
                if (!CsvValueConverter.TryConvert(
                        rawValue,
                        binding.Property.PropertyType,
                        document.Culture,
                        document.DateTimeFormats,
                        out object? converted,
                        out string? error))
                {
                    throw new CsvException(
                        $"Column '{document.Header[binding.ColumnIndex]}' cannot be assigned to " +
                        $"{typeof(T).Name}.{binding.Property.Name}: {error}");
                }

                binding.Property.SetValue(instance, converted, index: null);
            }

            yield return (T)instance;
        }
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

    private readonly record struct AutomaticMappingBinding(int ColumnIndex, PropertyInfo Property);
}
