#nullable enable

using System.Globalization;

namespace OfficeIMO.CSV;

/// <summary>
/// Represents a single row in a CSV document.
/// </summary>
public sealed class CsvRow
{
    private readonly CsvDocument _document;

    internal CsvRow(CsvDocument document, object?[] values)
    {
        _document = document;
        Values = values;
    }

    internal object?[] Values { get; set; }

    /// <summary>
    /// Gets or sets a field value by zero-based index.
    /// </summary>
    public object? this[int index]
    {
        get => Values[index];
        set => Values[index] = value;
    }

    /// <summary>
    /// Gets or sets a field value by column name.
    /// </summary>
    public object? this[string columnName]
    {
        get => Values[_document.GetColumnIndex(columnName)];
        set => Values[_document.GetColumnIndex(columnName)] = value;
    }

    /// <summary>
    /// Retrieves a typed value by column name.
    /// </summary>
    public T? Get<T>(string columnName) => Get<T>(_document.GetColumnIndex(columnName));

    /// <summary>
    /// Retrieves a typed value by index.
    /// </summary>
    public T? Get<T>(int index)
    {
        var value = Values[index];
        return CsvValueConverter.ConvertTo<T>(value, _document.Culture);
    }

    /// <summary>
    /// Returns the number of fields in the row.
    /// </summary>
    public int FieldCount => Values.Length;

    internal CsvRow CloneFor(CsvDocument document)
    {
        var copy = new object?[Values.Length];
        Array.Copy(Values, copy, Values.Length);
        return new CsvRow(document, copy);
    }

    internal CultureInfo Culture => _document.Culture;
}
