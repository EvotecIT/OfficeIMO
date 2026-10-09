#nullable enable
#if NET8_0_OR_GREATER
using System.Threading;
using OfficeIMO.Data;

namespace OfficeIMO.CSV;

/// <summary>Updates an accumulator from one transient CSV record.</summary>
/// <typeparam name="TAccumulator">Caller-owned accumulator state, including value types.</typeparam>
/// <param name="accumulator">State used exclusively by the current partition.</param>
/// <param name="record">Borrowed record whose field spans are valid only during this callback.</param>
/// <remarks>
/// Callbacks for different partitions can run concurrently. Modify only the supplied state and
/// do not retain the record or its borrowed spans. Reference-type state must be independent for
/// each call to the accumulator factory.
/// </remarks>
public delegate void CsvRecordAccumulator<TAccumulator>(ref TAccumulator accumulator, CsvRecord record);

/// <summary>Resolves headers once and builds a callback that updates CSV accumulator state.</summary>
/// <typeparam name="TAccumulator">Caller-owned accumulator state.</typeparam>
/// <param name="header">Resolved column names and ordinals.</param>
/// <returns>A callback that may run concurrently on independent states.</returns>
public delegate CsvRecordAccumulator<TAccumulator> CsvRecordAccumulatorBuilder<TAccumulator>(CsvRecordHeader header);

public sealed partial class CsvDocument
{
    /// <summary>Aggregates decoded CSV records into independent partition states and merges them in source order.</summary>
    /// <typeparam name="TAccumulator">Caller-owned accumulator state, including value types.</typeparam>
    /// <param name="text">Already-decoded CSV text.</param>
    /// <param name="createAccumulator">Creates an independent, neutral state for the result and each partition. May run concurrently.</param>
    /// <param name="accumulatorBuilder">Called once on the calling thread to resolve headers and build the record callback.</param>
    /// <param name="merge">Combines the result with a completed partition state. Called on the calling thread in source-partition order.</param>
    /// <param name="loadOptions">Optional CSV parsing and input-limit settings.</param>
    /// <param name="readerOptions">Optional reader schema settings. ParallelProcessing must be omitted.</param>
    /// <param name="parallelOptions">Optional degree and source partition size.</param>
    /// <param name="cancellationToken">Cancels parsing, record callbacks between records, and merging between partitions.</param>
    /// <returns>The accumulated result. Empty input returns the neutral state.</returns>
    /// <remarks>
    /// The factory must supply a neutral identity for merging. Merge must be associative; partition
    /// grouping can change with the degree and batch size. Merge order follows the source, so an
    /// associative operation need not be commutative. Records within each partition keep file order.
    /// Eligible text uses the existing partition parser and retains at most one state per active
    /// partition, without producing a result object or array entry for each row. Schema conversion,
    /// strict-width validation and unsupported parsing settings use the canonical sequential reader.
    /// Accumulator memory and resources belong to the caller; this method never disposes user states.
    /// Borrowed records and spans cannot outlive a callback. Cancellation does not interrupt a callback
    /// already running. All started workers finish before the method returns or throws.
    /// This method consumes decoded text and does not read or buffer a file.
    /// </remarks>
    public static TAccumulator AggregateTextRowsAsParallel<TAccumulator>(
        string text,
        Func<TAccumulator> createAccumulator,
        CsvRecordAccumulatorBuilder<TAccumulator> accumulatorBuilder,
        Func<TAccumulator, TAccumulator, TAccumulator> merge,
        CsvLoadOptions? loadOptions = null,
        CsvDataReaderOptions? readerOptions = null,
        ParallelRowMappingOptions? parallelOptions = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(createAccumulator);
        ArgumentNullException.ThrowIfNull(accumulatorBuilder);
        ArgumentNullException.ThrowIfNull(merge);
        if (readerOptions?.ParallelProcessing is not null)
        {
            throw new ArgumentException(
                "AggregateTextRowsAsParallel uses parallelOptions; CsvDataReaderOptions.ParallelProcessing must be omitted.",
                nameof(readerOptions));
        }
        var parallel = parallelOptions ?? new ParallelRowMappingOptions();
        int degree = parallel.GetDegreeOfParallelism();
        int batchSize = parallel.GetBatchSize(CsvParser.GetPreferredTextParallelBatchSize());
        var parsing = loadOptions?.Clone() ?? new CsvLoadOptions();
        CancellationToken loadCancellationToken = parsing.CancellationToken;
        cancellationToken.ThrowIfCancellationRequested();
        loadCancellationToken.ThrowIfCancellationRequested();
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, loadCancellationToken);
        parsing.CancellationToken = stop.Token;
        try
        {
            stop.Token.ThrowIfCancellationRequested();
            using var reader = (CsvDataReader)OpenTextDataReader(text, parsing, readerOptions);
            CsvRecordAccumulator<TAccumulator> accumulate = accumulatorBuilder(new CsvRecordHeader(reader))
                ?? throw new InvalidOperationException("The CSV accumulator builder returned null.");
            stop.Token.ThrowIfCancellationRequested();
            TAccumulator result = createAccumulator();
            // The reader owns schema conversion and original record error positions.
            bool useSourceRecords = readerOptions?.Schema is null && readerOptions?.InferSchema != true
                && parsing.ColumnCountMismatchPolicy != CsvColumnCountMismatchPolicy.Strict;
            if (degree > 1 && useSourceRecords &&
                reader.TryPrepareTextPartitioning(stop.Token, out var source, out int dataStart) &&
                TryCreateTextPartitions(text, dataStart, degree, batchSize, source!.Options,
                    stop.Token, out var partitions))
            {
                result = AggregateTextPartitions(text, source!, partitions!, degree,
                    createAccumulator, accumulate, merge, result, stop);
            }
            else
            {
                AccumulateReaderRecords(reader, batchSize, useSourceRecords, accumulate, ref result, stop.Token);
            }
            stop.Token.ThrowIfCancellationRequested();
            return result;
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            cancellationToken.ThrowIfCancellationRequested();
            throw;
        }
        catch (OperationCanceledException) when (loadCancellationToken.IsCancellationRequested)
        {
            loadCancellationToken.ThrowIfCancellationRequested();
            throw;
        }
    }

    private static void AccumulateReaderRecords<TAccumulator>(
        CsvDataReader reader,
        int batchSize,
        bool useSourceRecords,
        CsvRecordAccumulator<TAccumulator> accumulate,
        ref TAccumulator result,
        CancellationToken cancellationToken)
    {
        if (useSourceRecords)
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.TryReadCsvRecordBatch(batchSize, cancellationToken, out var batch)) break;
                if (batch is null) return;
                using (batch)
                {
                    while (batch.Read())
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        accumulate(ref result, new CsvRecord(batch));
                    }
                }
            }
        }
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!reader.Read()) break;
            cancellationToken.ThrowIfCancellationRequested();
            accumulate(ref result, new CsvRecord(reader));
        }
    }
}
#endif
