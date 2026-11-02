#nullable enable

namespace OfficeIMO.CSV;

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
