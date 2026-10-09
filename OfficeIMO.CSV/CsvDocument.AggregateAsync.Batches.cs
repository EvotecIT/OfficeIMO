#nullable enable
#if NET8_0_OR_GREATER
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    private static async Task<TAccumulator> AggregateReaderTextBatchesAsync<TAccumulator>(
        CsvDataReader reader, int degree, int batchSize, Func<TAccumulator> createAccumulator,
        CsvRecordAccumulator<TAccumulator> accumulate, Func<TAccumulator, TAccumulator, TAccumulator> merge,
        TAccumulator result, CancellationTokenSource stop)
    {
        var pending = new Queue<Task<TAccumulator>>(degree);
        var failure = new CsvAggregateFailure();
        try
        {
            bool reachedEnd = false;
            while (!reachedEnd || pending.Count > 0)
            {
                stop.Token.ThrowIfCancellationRequested();
                while (!reachedEnd && pending.Count < degree)
                {
                    var capture = await reader.ReadRecordTextBatchAsync(batchSize, stop.Token).ConfigureAwait(false);
                    reachedEnd = capture.ReachedEnd;
                    if (capture.Batch.Count == 0) { capture.Batch.Dispose(); break; }
                    try
                    {
                        pending.Enqueue(Task.Factory.StartNew(
                            () => AccumulateRecordTextBatch(capture.Batch, createAccumulator, accumulate, stop, failure),
                            CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default));
                    }
                    catch { capture.Batch.Dispose(); throw; }
                }
                if (pending.Count == 0) break;
                TAccumulator state = await pending.Dequeue().ConfigureAwait(false);
                stop.Token.ThrowIfCancellationRequested();
                result = merge(result, state);
            }
            return result;
        }
        catch (Exception exception)
        {
            failure.Capture(exception, stop.Token);
            stop.Cancel();
            while (pending.Count > 0)
            {
                try { await pending.Dequeue().ConfigureAwait(false); }
                catch { } // Each worker captured its original failure before requesting stop.
            }
            failure.RethrowOriginal();
            ExceptionDispatchInfo.Capture(exception).Throw();
            throw;
        }
    }

    private static TAccumulator AccumulateRecordTextBatch<TAccumulator>(
        CsvRecordTextBatch batch, Func<TAccumulator> createAccumulator,
        CsvRecordAccumulator<TAccumulator> accumulate, CancellationTokenSource stop, CsvAggregateFailure failure)
    {
        using (batch)
        {
            try
            {
                stop.Token.ThrowIfCancellationRequested();
                TAccumulator state = createAccumulator();
                while (batch.Read())
                {
                    stop.Token.ThrowIfCancellationRequested();
                    accumulate(ref state, new CsvRecord(batch));
                }
                stop.Token.ThrowIfCancellationRequested();
                return state;
            }
            catch (Exception exception) { failure.Capture(exception, stop.Token); stop.Cancel(); throw; }
        }
    }
}
#endif
