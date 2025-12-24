#nullable enable

using System.Globalization;

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

    T Apply(T instance, object? rawValue, CultureInfo culture);
}

internal sealed class CsvMappingEntry<T, TValue> : ICsvMappingEntry<T>
{
    public CsvMappingEntry(string columnName, Func<T, TValue, T> assign)
    {
        ColumnName = columnName;
        _assign = assign;
    }

    public string ColumnName { get; }

    public T Apply(T instance, object? rawValue, CultureInfo culture)
    {
        var value = CsvValueConverter.ConvertTo<TValue>(rawValue, culture);
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
                instance = binding.Entry.Apply(instance, rawValue, document.Culture);
            }

            yield return instance;
        }
    }

    private readonly record struct MappingBinding<T>(int ColumnIndex, ICsvMappingEntry<T> Entry);
}
