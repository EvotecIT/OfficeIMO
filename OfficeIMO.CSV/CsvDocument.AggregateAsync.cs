#nullable enable
#if NET8_0_OR_GREATER
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>Aggregates CSV file records with incremental asynchronous I/O and bounded worker-owned states.</summary>
    /// <typeparam name="TAccumulator">Caller-owned accumulator state, including value types.</typeparam>
    /// <param name="path">Source CSV file.</param>
    /// <param name="createAccumulator">Creates an independent neutral state for the result and each batch. May run concurrently.</param>
    /// <param name="accumulatorBuilder">Called once after resolving the header to build the record callback.</param>
    /// <param name="merge">Associatively combines completed states in source-batch order.</param>
    /// <param name="loadOptions">Optional parsing, encoding, compression, cancellation and input-limit settings.</param>
    /// <param name="readerOptions">Optional schema settings. ParallelProcessing must be omitted.</param>
    /// <param name="parallelOptions">Optional degree and maximum records per batch.</param>
    /// <param name="cancellationToken">Cancels input, processing between records, and merging between batches.</param>
    /// <returns>The accumulated result. Empty input returns the neutral state.</returns>
    /// <remarks>
    /// Reads incrementally without loading the whole file. The existing parser materializes decoded
    /// field strings; bounded pooled batches retain their references and null/missing metadata while
    /// the reader advances. At most the configured degree of batches are retained, with one state
    /// per batch and no mapped result per row. Callback spans remain valid only during the callback.
    /// Explicit or inferred schema conversion uses native asynchronous sequential consumption.
    /// Factory states must be independent and neutral for associative merge; merge need not be
    /// commutative. Record callbacks and factories may run concurrently, while merge calls are
    /// sequential in source order. No callback has thread affinity. User states are never disposed.
    /// Cancellation cannot interrupt a running callback; all started workers finish before return
    /// or failure. The file reader is closed on success and failure.
    /// </remarks>
    public static Task<TAccumulator> AggregateRowsAsParallelAsync<TAccumulator>(
        string path, Func<TAccumulator> createAccumulator,
        CsvRecordAccumulatorBuilder<TAccumulator> accumulatorBuilder,
        Func<TAccumulator, TAccumulator, TAccumulator> merge,
        CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null,
        ParallelRowMappingOptions? parallelOptions = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        return AggregateRowsAsParallelAsyncCore(
            (options, token) => OpenDataReaderAsync(path, options, readerOptions, token),
            createAccumulator, accumulatorBuilder, merge, loadOptions, readerOptions, parallelOptions, cancellationToken);
    }

    /// <summary>Aggregates CSV records from the stream's current position using incremental asynchronous I/O.</summary>
    /// <typeparam name="TAccumulator">Caller-owned accumulator state, including value types.</typeparam>
    /// <param name="stream">Readable source. Its position advances and it remains open on success and failure.</param>
    /// <param name="createAccumulator">Creates independent neutral result and batch states. May run concurrently.</param>
    /// <param name="accumulatorBuilder">Called once after resolving the header to build the record callback.</param>
    /// <param name="merge">Associatively combines completed states in source-batch order.</param>
    /// <param name="loadOptions">Optional parsing, encoding, compression, cancellation and input-limit settings.</param>
    /// <param name="readerOptions">Optional schema settings. ParallelProcessing must be omitted.</param>
    /// <param name="parallelOptions">Optional degree and maximum records per batch.</param>
    /// <param name="cancellationToken">Cancels input, processing between records, and merging between batches.</param>
    /// <returns>The accumulated result. Empty input returns the neutral state.</returns>
    /// <remarks>
    /// Uses the same bounded decoded-field batches, callback lifetime, independent state ownership,
    /// source-ordered associative merge and schema fallback as the file overload. It performs actual
    /// asynchronous reads and never snapshots the whole stream. All started workers finish before
    /// return or failure; the caller retains ownership of the stream.
    /// </remarks>
    public static Task<TAccumulator> AggregateRowsAsParallelAsync<TAccumulator>(
        Stream stream, Func<TAccumulator> createAccumulator,
        CsvRecordAccumulatorBuilder<TAccumulator> accumulatorBuilder,
        Func<TAccumulator, TAccumulator, TAccumulator> merge,
        CsvLoadOptions? loadOptions = null, CsvDataReaderOptions? readerOptions = null,
        ParallelRowMappingOptions? parallelOptions = null, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(stream);
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
        return AggregateRowsAsParallelAsyncCore(
            (options, token) => OpenDataReaderAsync(stream, options, readerOptions, token),
            createAccumulator, accumulatorBuilder, merge, loadOptions, readerOptions, parallelOptions, cancellationToken);
    }

    private static async Task<TAccumulator> AggregateRowsAsParallelAsyncCore<TAccumulator>(
        Func<CsvLoadOptions, CancellationToken, Task<DbDataReader>> openReader,
        Func<TAccumulator> createAccumulator, CsvRecordAccumulatorBuilder<TAccumulator> accumulatorBuilder,
        Func<TAccumulator, TAccumulator, TAccumulator> merge, CsvLoadOptions? loadOptions,
        CsvDataReaderOptions? readerOptions, ParallelRowMappingOptions? parallelOptions, CancellationToken token)
    {
        ArgumentNullException.ThrowIfNull(createAccumulator);
        ArgumentNullException.ThrowIfNull(accumulatorBuilder);
        ArgumentNullException.ThrowIfNull(merge);
        if (readerOptions?.ParallelProcessing is not null)
            throw new ArgumentException("AggregateRowsAsParallelAsync uses parallelOptions; CsvDataReaderOptions.ParallelProcessing must be omitted.", nameof(readerOptions));
        var parallel = parallelOptions ?? new ParallelRowMappingOptions();
        int degree = parallel.GetDegreeOfParallelism(), batchSize = parallel.GetBatchSize(256);
        var parsing = loadOptions?.Clone() ?? new CsvLoadOptions();
        CancellationToken loadToken = parsing.CancellationToken;
        token.ThrowIfCancellationRequested();
        loadToken.ThrowIfCancellationRequested();
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(token, loadToken);
        parsing.CancellationToken = stop.Token;
        try
        {
            using var reader = (CsvDataReader)await openReader(parsing, stop.Token).ConfigureAwait(false);
            var accumulate = accumulatorBuilder(new CsvRecordHeader(reader))
                ?? throw new InvalidOperationException("The CSV accumulator builder returned null.");
            stop.Token.ThrowIfCancellationRequested();
            TAccumulator result = createAccumulator();
            if (degree > 1 && readerOptions?.Schema is null && readerOptions?.InferSchema != true)
                result = await AggregateReaderTextBatchesAsync(reader, degree, batchSize,
                    createAccumulator, accumulate, merge, result, stop).ConfigureAwait(false);
            else
            {
                while (await reader.ReadAsync(stop.Token).ConfigureAwait(false))
                {
                    stop.Token.ThrowIfCancellationRequested();
                    accumulate(ref result, new CsvRecord(reader));
                }
            }
            stop.Token.ThrowIfCancellationRequested();
            return result;
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested)
        {
            token.ThrowIfCancellationRequested();
            throw;
        }
        catch (OperationCanceledException) when (loadToken.IsCancellationRequested)
        {
            loadToken.ThrowIfCancellationRequested();
            throw;
        }
    }
}
#endif
