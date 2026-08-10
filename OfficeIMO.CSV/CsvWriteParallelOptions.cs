#nullable enable

namespace OfficeIMO.CSV;

/// <summary>
/// Controls bounded, ordered parallel formatting for streaming CSV writes.
/// </summary>
public sealed class CsvWriteParallelOptions
{
    /// <summary>
    /// Gets or sets the maximum number of row-formatting workers. When omitted,
    /// the writer uses up to four workers.
    /// </summary>
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>
    /// Gets or sets the number of rows in each pipeline batch. Default is 4,096 rows.
    /// </summary>
    public int BatchSize { get; set; } = 4096;

    internal int GetDegreeOfParallelism()
    {
        if (MaxDegreeOfParallelism is int configured && configured <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaxDegreeOfParallelism),
                "MaxDegreeOfParallelism must be greater than zero when specified.");
        }

        return MaxDegreeOfParallelism ?? Math.Min(4, Math.Max(1, Environment.ProcessorCount));
    }

    internal int GetBatchSize()
    {
        if (BatchSize <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(BatchSize), "BatchSize must be greater than zero.");
        }

        return BatchSize;
    }
}
