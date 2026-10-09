#nullable enable
#if NET8_0_OR_GREATER
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private static TAccumulator AggregateTextPartitions<TAccumulator>(
        string text,
        CsvParser.CsvTextDataReaderRowSource preparedSource,
        CsvTextPartition[] partitions,
        int degree,
        Func<TAccumulator> createAccumulator,
        CsvRecordAccumulator<TAccumulator> accumulate,
        Func<TAccumulator, TAccumulator, TAccumulator> merge,
        TAccumulator result,
        CancellationTokenSource stop)
    {
        var states = new TAccumulator[Math.Min(degree, partitions.Length)];
        var failure = new CsvAggregateFailure();
        for (int waveStart = 0; waveStart < partitions.Length; waveStart += states.Length)
        {
            stop.Token.ThrowIfCancellationRequested();
            int count = Math.Min(states.Length, partitions.Length - waveStart);
            try
            {
                try
                {
                    int capturedStart = waveStart;
                    Parallel.For(0, count,
                        new ParallelOptions { MaxDegreeOfParallelism = count, CancellationToken = stop.Token },
                        index => {
                            try
                            {
                                stop.Token.ThrowIfCancellationRequested();
                                TAccumulator state = createAccumulator();
                                AccumulateTextPartition(text, preparedSource.Options, preparedSource.SourceColumnCount,
                                    partitions[capturedStart + index], accumulate, ref state, stop.Token);
                                states[index] = state;
                            }
                            catch (Exception exception)
                            {
                                failure.Capture(exception, stop.Token);
                                stop.Cancel();
                                throw;
                            }
                        });
                }
                catch (Exception exception)
                {
                    failure.RethrowOriginal();
                    if (exception is AggregateException aggregate)
                    {
                        var errors = aggregate.Flatten().InnerExceptions;
                        Exception first = errors.FirstOrDefault(static error => error is not OperationCanceledException)
                            ?? errors.First();
                        ExceptionDispatchInfo.Capture(first).Throw();
                    }
                    throw;
                }
                for (int index = 0; index < count; index++)
                {
                    stop.Token.ThrowIfCancellationRequested();
                    result = merge(result, states[index]);
                }
            }
            finally
            {
                // Parallel.For joins all workers before returning or propagating a failure.
                // Release only our references; user states retain caller-defined ownership.
                Array.Clear(states, 0, count);
            }
        }
        return result;
    }

    private static void AccumulateTextPartition<TAccumulator>(
        string text,
        CsvLoadOptions options,
        int columnCount,
        CsvTextPartition partition,
        CsvRecordAccumulator<TAccumulator> accumulate,
        ref TAccumulator state,
        CancellationToken cancellationToken)
    {
        using var source = new CsvParser.CsvTextDataReaderRowSource(
            text, options, recordsToSkip: 0, columnCount, partition.Start, partition.End);
        int rows = 0;
        while (source.Read(cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (rows == partition.RowCount)
                throw new InvalidDataException("Partitioned CSV record counting did not match parser output.");
            accumulate(ref state, new CsvRecord(source));
            rows++;
        }
        cancellationToken.ThrowIfCancellationRequested();
        if (rows != partition.RowCount)
            throw new InvalidDataException("Partitioned CSV record counting did not match parser output.");
    }
}
#endif
