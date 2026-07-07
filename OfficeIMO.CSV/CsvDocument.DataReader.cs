#nullable enable

using System.Data.Common;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    // Large explicit schema samples are faster when streamed twice than kept live through reader traversal.
    private const int StreamingInferredReaderBufferLimit = 1000;

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

        if (options.Schema is null &&
            _schema is null &&
            options.InferSchema &&
            options.SchemaSampleSize <= StreamingInferredReaderBufferLimit &&
            _mode == CsvLoadMode.Stream &&
            _streamingSource is not null)
        {
            return CreateStreamingInferredDataReader(options.SchemaSampleSize);
        }

        var schema = options.Schema ?? _schema ?? (options.InferSchema ? InferSchema(options.SchemaSampleSize) : null);
        var schemaColumns = schema?.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
        var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
        return new CsvDataReader(columns, EnumerateRawRows(), _culture, _dateTimeFormats);
    }

    private CsvDataReader CreateStreamingInferredDataReader(int schemaSampleSize)
    {
        var rows = EnumerateRawRows().GetEnumerator();
        try
        {
            var sampledRows = new List<object?[]>(Math.Min(schemaSampleSize, 4096));
            var schema = InferSchema(rows, schemaSampleSize, sampledRows);
            var schemaColumns = schema.Columns.ToDictionary(column => column.Name, StringComparer.OrdinalIgnoreCase);
            var columns = CsvDataProjectionBuilder.Create(_header, schemaColumns);
            return new CsvDataReader(columns, EnumerateSampledThenRemainingRows(sampledRows, rows), _culture, _dateTimeFormats);
        }
        catch
        {
            rows.Dispose();
            throw;
        }
    }

    private static IEnumerable<object?[]> EnumerateSampledThenRemainingRows(IReadOnlyList<object?[]> sampledRows, IEnumerator<object?[]> remainingRows)
    {
        try
        {
            for (var i = 0; i < sampledRows.Count; i++)
            {
                yield return sampledRows[i];
            }

            while (remainingRows.MoveNext())
            {
                yield return remainingRows.Current;
            }
        }
        finally
        {
            remainingRows.Dispose();
        }
    }
}
