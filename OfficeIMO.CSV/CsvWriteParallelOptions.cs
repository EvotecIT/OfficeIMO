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

    /// <summary>
    /// Gets or sets the maximum number of field slots retained in each row batch.
    /// The writer retains at most two batches. Default is 1,048,576 field slots per batch.
    /// </summary>
    public int MaximumBufferedCellsPerBatch { get; set; } = 1_048_576;

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

    internal int GetBatchSize(int fieldCount)
    {
        int requestedBatchSize = GetBatchSize();
        if (MaximumBufferedCellsPerBatch <= 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(MaximumBufferedCellsPerBatch),
                "MaximumBufferedCellsPerBatch must be greater than zero.");
        }
        if (fieldCount <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(fieldCount), "Field count must be greater than zero.");
        }
        if (fieldCount > MaximumBufferedCellsPerBatch)
        {
            throw new InvalidOperationException(
                $"Data reader exposes {fieldCount} fields, exceeding the configured per-batch cell budget of {MaximumBufferedCellsPerBatch}.");
        }

        return Math.Min(requestedBatchSize, MaximumBufferedCellsPerBatch / fieldCount);
    }
}
