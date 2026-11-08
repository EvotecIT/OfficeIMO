using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data.Common;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Data;

public static partial class ParallelRowMappingExtensions {
    private static IEnumerable<T> EnumerateReaderBatches<T>(
        DbDataReader sourceReader,
        IDataReaderParallelBatchSource batchSource,
        int preferredBatchSize,
        int degreeOfParallelism,
        Func<DbDataReader, IEnumerable<T>> mapReader,
        Func<IEnumerable<T>> mapGenericRemainder,
        CancellationToken cancellationToken) {
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        using var results = new ReaderResultPool<T>(degreeOfParallelism);
        var pending = new Queue<Task<T[]>>(degreeOfParallelism);
        bool useGenericRemainder = false;
        ExceptionDispatchInfo? sourceError = null;
        try {
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                DbDataReader? batchReader;
                try {
                    if (!batchSource.TryReadParallelBatch(
                            preferredBatchSize,
                            cancellationToken,
                            out batchReader)) {
                        useGenericRemainder = true;
                        break;
                    }
                } catch (Exception exception) when (!cancellationToken.IsCancellationRequested) {
                    sourceError = ExceptionDispatchInfo.Capture(exception);
                    break;
                }
                if (batchReader is null) break;

                pending.Enqueue(Task.Factory.StartNew(
                    () => MapReaderBatch(batchReader, mapReader, results, stop.Token),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));

                if (pending.Count < degreeOfParallelism) continue;
                foreach (T row in AwaitNextReaderBatch(pending, results, stop, cancellationToken)) yield return row;
            }

            while (pending.Count > 0) {
                foreach (T row in AwaitNextReaderBatch(pending, results, stop, cancellationToken)) yield return row;
            }

            if (useGenericRemainder) {
                foreach (T row in mapGenericRemainder()) {
                    yield return row;
                }
            }
            sourceError?.Throw();
        } finally {
            stop.Cancel();
            while (pending.Count > 0) {
                Task<T[]> task = pending.Dequeue();
                try {
                    results.Return(task.GetAwaiter().GetResult());
                } catch {
                    // Observe remaining canceled/faulted work after the ordered exception is propagated.
                }
            }
        }
    }

    private static T[] MapReaderBatch<T>(
        DbDataReader reader,
        Func<DbDataReader, IEnumerable<T>> mapReader,
        ReaderResultPool<T> results,
        CancellationToken cancellationToken) {
        T[]? rented = null;
        T[]? rows = null;
        bool completed = false;
        try {
            using (reader) {
                cancellationToken.ThrowIfCancellationRequested();
                int expectedCount = reader is IDataReaderParallelBatchInfo info ? info.ParallelBatchRowCount : 0;
                rented = results.Rent(expectedCount);
                rows = MaterializeRows(mapReader(reader), cancellationToken, expectedCount, rented);
            }
            completed = true;
            return rows;
        } finally {
            // A resized result owns its copied entries; a failed worker publishes no result.
            if (rented is not null && (!completed || !ReferenceEquals(rows, rented)))
                Array.Clear(rented, 0, rented.Length);
            if (!completed && rows is not null && !ReferenceEquals(rows, rented))
                Array.Clear(rows, 0, rows.Length);
        }
    }

    private static T[] MaterializeRows<T>(
        IEnumerable<T> rows,
        CancellationToken cancellationToken,
        int expectedCount,
        T[]? rented) {
        if (expectedCount > 0) {
            T[] exact = rented ?? new T[expectedCount];
            int count = 0;
            try {
                foreach (T row in rows) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (count == exact.Length) Array.Resize(ref exact, checked(exact.Length * 2));
                    exact[count++] = row;
                }
                if (count != exact.Length) Array.Resize(ref exact, count);
                return exact;
            } catch {
                Array.Clear(exact, 0, exact.Length);
                throw;
            }
        }
        var materialized = new List<T>();
        foreach (T row in rows) {
            cancellationToken.ThrowIfCancellationRequested();
            materialized.Add(row);
        }
        return materialized.ToArray();
    }

    private static IEnumerable<T> AwaitNextReaderBatch<T>(
        Queue<Task<T[]>> pending,
        ReaderResultPool<T> results,
        CancellationTokenSource stop,
        CancellationToken cancellationToken) {
        T[] rows;
        try {
            rows = pending.Dequeue().GetAwaiter().GetResult();
        } catch {
            stop.Cancel();
            throw;
        }
        try {
            for (int index = 0; index < rows.Length; index++) {
                if (cancellationToken.CanBeCanceled) cancellationToken.ThrowIfCancellationRequested();
                yield return rows[index];
            }
        } finally {
            results.Return(rows);
        }
    }

    /// <summary>Shares cleared result arrays only within one bounded, ordered enumeration.</summary>
    private sealed class ReaderResultPool<T> : IDisposable {
        private readonly ConcurrentQueue<T[]> _available = new();
        private readonly int _capacity;

        internal ReaderResultPool(int capacity) => _capacity = capacity;

        internal T[]? Rent(int expectedCount) {
            if (expectedCount <= 0) return null;
            if (_available.TryDequeue(out T[]? rows) && rows.Length == expectedCount) return rows;
            return new T[expectedCount];
        }

        internal void Return(T[] rows) {
            Array.Clear(rows, 0, rows.Length);
            // Sources without a row-count hint retain their ordinary materialization fallback.
            if (rows.Length != 0 && _available.Count < _capacity) _available.Enqueue(rows);
        }

        public void Dispose() {
            // All workers have joined, and every queued result has already been cleared.
            while (_available.TryDequeue(out _)) { }
        }
    }
}
