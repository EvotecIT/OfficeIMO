#nullable enable

using System.Data.Common;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Creates a forward-only data reader over the document rows.
    /// </summary>
    /// <param name="options">Reader projection options. When omitted, all columns are emitted as strings.</param>
    /// <returns>A data reader suitable for DataTable loading and provider bulk-copy APIs.</returns>
    public CsvDataReader CreateDataReader(CsvDataReaderOptions? options = null)
    {
        options ??= new CsvDataReaderOptions();
        if (options.SchemaSampleSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "Schema sample size must be greater than zero.");
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var schemaColumns = schema?.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
        var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
        return new CsvDataReader(columns, EnumerateRawRows(), _culture, _dateTimeFormats);
    }
}
