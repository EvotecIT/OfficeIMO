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
        var pending = new Queue<CsvAggregateBatchSlot<TAccumulator>>(degree);
        var slots = new List<CsvAggregateBatchSlot<TAccumulator>>(degree);
        CsvAggregateBatchSlot<TAccumulator>? available = null;
        var failure = new CsvAggregateFailure();
        try
        {
            bool reachedEnd = false;
            while (!reachedEnd || pending.Count > 0)
            {
                stop.Token.ThrowIfCancellationRequested();
                while (!reachedEnd && pending.Count < degree)
                {
                    CsvAggregateBatchSlot<TAccumulator> slot;
                    if (available is null)
                    {
                        slot = new CsvAggregateBatchSlot<TAccumulator>(reader, batchSize,
                            createAccumulator, accumulate, stop, failure);
                        slots.Add(slot);
                    }
                    else { slot = available; available = null; }
                    reachedEnd = await reader.ReadRecordTextBatchAsync(slot.Batch, stop.Token).ConfigureAwait(false);
                    if (slot.Batch.Count == 0) break;
                    slot.Start();
                    pending.Enqueue(slot);
                }
                if (pending.Count == 0) break;
                var completed = pending.Dequeue();
                TAccumulator state = await completed.Work!.ConfigureAwait(false);
                stop.Token.ThrowIfCancellationRequested();
                result = merge(result, state);
                completed.Work = null;
                available = completed;
            }
            return result;
        }
        catch (Exception exception)
        {
            failure.Capture(exception, stop.Token);
            stop.Cancel();
            while (pending.Count > 0)
            {
                try { await pending.Dequeue().Work!.ConfigureAwait(false); }
                catch { } // Each worker captured its original failure before requesting stop.
            }
            failure.RethrowOriginal();
            ExceptionDispatchInfo.Capture(exception).Throw();
            throw;
        }
        finally
        {
            // A slot remains owned here until its worker has stopped, even if dispatch or
            // cancellation throws before it enters the pending queue.
            foreach (var slot in slots)
            {
                if (slot.Work is { } work)
                {
                    try { await work.ConfigureAwait(false); }
                    catch { }
                }
            }
            foreach (var slot in slots) slot.Batch.Dispose();
        }
    }

    private static TAccumulator AccumulateRecordTextBatch<TAccumulator>(
        CsvRecordTextBatch batch, Func<TAccumulator> createAccumulator,
        CsvRecordAccumulator<TAccumulator> accumulate, CancellationTokenSource stop, CsvAggregateFailure failure)
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

    // Slots and their buffers are bounded by the degree; tasks still capture the producer's
    // ExecutionContext separately for every original source batch.
    private sealed class CsvAggregateBatchSlot<TAccumulator>
    {
        private static readonly Func<object?, TAccumulator> Callback = static state =>
        {
            var slot = (CsvAggregateBatchSlot<TAccumulator>)state!;
            return AccumulateRecordTextBatch(slot.Batch, slot._createAccumulator,
                slot._accumulate, slot._stop, slot._failure);
        };
        private readonly Func<TAccumulator> _createAccumulator;
        private readonly CsvRecordAccumulator<TAccumulator> _accumulate;
        private readonly CancellationTokenSource _stop;
        private readonly CsvAggregateFailure _failure;

        internal CsvAggregateBatchSlot(CsvDataReader reader, int batchSize,
            Func<TAccumulator> createAccumulator, CsvRecordAccumulator<TAccumulator> accumulate,
            CancellationTokenSource stop, CsvAggregateFailure failure)
        {
            _createAccumulator = createAccumulator;
            _accumulate = accumulate;
            _stop = stop;
            _failure = failure;
            Batch = reader.CreateRecordTextBatch(batchSize);
        }

        internal CsvRecordTextBatch Batch { get; }
        internal Task<TAccumulator>? Work { get; set; }
        internal void Start() => Work = Task.Factory.StartNew(Callback, this,
            CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
    }
}
#endif
