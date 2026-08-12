#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls how a <see cref="CsvDocument"/> is exposed as a forward-only data reader.
/// </summary>
public sealed class CsvDataReaderOptions
{
    /// <summary>
    /// Gets or sets an explicit schema used for typed reader columns.
    /// </summary>
    public CsvSchema? Schema { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether schema should be inferred before creating reader columns.
    /// </summary>
    public bool InferSchema { get; set; }

    /// <summary>
    /// Gets or sets the maximum row count inspected when <see cref="InferSchema"/> is enabled.
    /// </summary>
    public int SchemaSampleSize { get; set; } = 1000;

    /// <summary>
    /// Gets or sets bounded, ordered parallel processing for typed reader values.
    /// When omitted, the reader remains fully sequential.
    /// </summary>
    /// <remarks>
    /// CSV records are parsed by one source owner and typed values are projected by bounded
    /// worker batches. Row order and the single-consumer <see cref="System.Data.Common.DbDataReader"/>
    /// contract are preserved, making the reader suitable for provider bulk-copy APIs. Schema
    /// converters can be invoked concurrently and must be thread-safe; use sequential processing
    /// when a converter depends on mutable single-threaded state.
    /// </remarks>
    public CsvDataReaderParallelOptions? ParallelProcessing { get; set; }
}

/// <summary>
/// Controls bounded, ordered parallel processing for a CSV data reader.
/// </summary>
public sealed class CsvDataReaderParallelOptions
{
    /// <summary>
    /// Gets or sets the maximum number of batches projected concurrently.
    /// A null value uses up to four workers, bounded by <see cref="Environment.ProcessorCount"/>.
    /// </summary>
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>
    /// Gets or sets the number of rows in each worker batch.
    /// A null value lets the reader choose its tuned bounded default.
    /// </summary>
    public int? BatchSize { get; set; }

    internal int GetDegreeOfParallelism()
    {
        if (MaxDegreeOfParallelism is int configured && configured <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaxDegreeOfParallelism),
                "MaxDegreeOfParallelism must be greater than zero when specified.");
        }

        if (BatchSize is int batchSize && batchSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(BatchSize), "BatchSize must be greater than zero.");
        }

        return MaxDegreeOfParallelism ?? Math.Min(4, Math.Max(1, Environment.ProcessorCount));
    }

    internal int GetBatchSize(int sourceDefault) => BatchSize ?? sourceDefault;
}
