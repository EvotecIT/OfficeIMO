#nullable enable

using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.ExceptionServices;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Data;

/// <summary>Controls bounded, ordered parallel projection from a forward-only data reader.</summary>
public sealed class ParallelRowMappingOptions {
    /// <summary>
    /// Gets or sets the maximum number of row batches projected concurrently.
    /// A null value uses <see cref="Environment.ProcessorCount"/>.
    /// </summary>
    public int? MaxDegreeOfParallelism { get; set; }

    /// <summary>
    /// Gets or sets the number of source rows copied into each worker batch.
    /// A null value lets the source choose its tuned bounded default.
    /// </summary>
    public int? BatchSize { get; set; }

    internal int GetDegreeOfParallelism() {
        if (MaxDegreeOfParallelism is int configured && configured <= 0) {
            throw new ArgumentOutOfRangeException(
                nameof(MaxDegreeOfParallelism),
                "MaxDegreeOfParallelism must be greater than zero when specified.");
        }
        if (BatchSize is int batchSize && batchSize <= 0) {
            throw new ArgumentOutOfRangeException(nameof(BatchSize), "BatchSize must be greater than zero.");
        }
        return MaxDegreeOfParallelism ?? Math.Max(1, Environment.ProcessorCount);
    }

    internal int GetBatchSize(int sourceDefault) => BatchSize ?? sourceDefault;
}

/// <summary>Ordered parallel typed projections over forward-only data readers.</summary>
public static class ParallelRowMappingExtensions {
    internal static IEnumerable<T> RowsAsParallelValues<T>(
        this DbDataReader reader,
        Func<object?[], T> map,
        ParallelRowMappingOptions options,
        CancellationToken cancellationToken) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (map is null) throw new ArgumentNullException(nameof(map));
        if (options is null) throw new ArgumentNullException(nameof(options));
        int degreeOfParallelism = options.GetDegreeOfParallelism();
        if (reader.FieldCount == 0) yield break;
        foreach (T row in EnumerateBatches(
                     reader,
                     options.GetBatchSize(128),
                     degreeOfParallelism,
                     map,
                     captureValues: null,
                     cancellationToken)) {
            yield return row;
        }
    }

    /// <summary>
    /// Projects rows in bounded parallel batches by matching columns to writable public properties.
    /// Only the calling thread reads the source reader; worker tasks receive independent row snapshots.
    /// Results retain source order.
    /// </summary>
    /// <remarks>
    /// Enumerate the returned sequence while the reader remains open. A degree of parallelism of one
    /// uses the ordinary sequential <see cref="DataReaderMappingExtensions.RowsAs{T}(DbDataReader)"/> path.
    /// </remarks>
    public static IEnumerable<T> RowsAsParallel<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        this DbDataReader reader,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        return EnumerateAutomatic<T>(reader, options, cancellationToken);
    }

    /// <summary>
    /// Projects rows in bounded parallel batches using explicit, AOT-friendly column assignments.
    /// Results retain source order.
    /// </summary>
    public static IEnumerable<T> RowsAsParallel<T>(
        this DbDataReader reader,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) where T : new() {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        return EnumerateExplicit(reader, configure, options, cancellationToken);
    }

    /// <summary>
    /// Projects rows in bounded parallel batches with a caller-supplied factory.
    /// Results retain source order.
    /// </summary>
    /// <remarks>
    /// The factory may run concurrently and must not mutate shared state. The supplied
    /// <see cref="IDataRecord"/> represents only the current factory call and must not be retained.
    /// Parallel factory execution requires a reader that can provide independent batch readers;
    /// other readers preserve their native <see cref="IDataRecord"/> behavior on the calling thread.
    /// </remarks>
    public static IEnumerable<T> RowsAsParallel<T>(
        this DbDataReader reader,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        if (factory is null) throw new ArgumentNullException(nameof(factory));
        return EnumerateFactory(reader, factory, options, cancellationToken);
    }

    private static IEnumerable<T> EnumerateAutomatic<
        [DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicProperties)] T>(
        DbDataReader reader,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) where T : new() {
        options ??= new ParallelRowMappingOptions();
        int degreeOfParallelism = options.GetDegreeOfParallelism();
        if (degreeOfParallelism == 1) {
            foreach (T row in EnumerateSequential(reader.RowsAs<T>(), cancellationToken)) {
                yield return row;
            }
            yield break;
        }
        if (reader.FieldCount == 0) yield break;

        DataReaderMappingExtensions.GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out bool requireAllColumnsMapped,
            out DataMappingErrorValuePolicy errorValuePolicy);
        AutomaticRowMappingPlan<T> plan = AutomaticRowMappingPlan<T>.Create(
            DataReaderMappingExtensions.GetHeaders(reader),
            requireAllColumnsMapped);
        int genericBatchSize = options.GetBatchSize(128);

        if (reader is IDataReaderParallelBatchSource { CanReadParallelBatches: true } batchSource) {
            foreach (T row in EnumerateReaderBatches(
                         reader,
                         batchSource,
                         options.GetBatchSize(batchSource.PreferredParallelBatchSize),
                         degreeOfParallelism,
                         static batchReader => batchReader.RowsAs<T>(),
                         () => EnumerateBatches(
                             reader,
                             genericBatchSize,
                             degreeOfParallelism,
                             values => plan.MapValues(values, culture, dateTimeFormats, typeConverter, errorValuePolicy),
                             typeConverter is null ? plan.CaptureReaderValues : null,
                             cancellationToken),
                         cancellationToken)) {
                yield return row;
            }
            yield break;
        }

        foreach (T row in EnumerateBatches(
                     reader,
                     genericBatchSize,
                     degreeOfParallelism,
                     values => plan.MapValues(values, culture, dateTimeFormats, typeConverter, errorValuePolicy),
                     typeConverter is null ? plan.CaptureReaderValues : null,
                     cancellationToken)) {
            yield return row;
        }
    }

    private static IEnumerable<T> EnumerateExplicit<T>(
        DbDataReader reader,
        Action<RowMapper<T>> configure,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) where T : new() {
        options ??= new ParallelRowMappingOptions();
        int degreeOfParallelism = options.GetDegreeOfParallelism();
        if (degreeOfParallelism == 1) {
            foreach (T row in EnumerateSequential(reader.RowsAs(configure), cancellationToken)) {
                yield return row;
            }
            yield break;
        }
        if (reader.FieldCount == 0) yield break;

        ExplicitRowMappingPlan<T> plan = ExplicitRowMappingPlan<T>.Create(
            DataReaderMappingExtensions.GetHeaders(reader),
            configure);
        if (plan.IsEmpty) yield break;
        DataReaderMappingExtensions.GetConversionOptions(
            reader,
            out CultureInfo culture,
            out IReadOnlyList<string>? dateTimeFormats,
            out Func<object, Type, CultureInfo, (bool ok, object? value)>? typeConverter,
            out _,
            out DataMappingErrorValuePolicy errorValuePolicy);
        int genericBatchSize = options.GetBatchSize(128);

        if (reader is IDataReaderParallelBatchSource { CanReadParallelBatches: true } batchSource) {
            foreach (T row in EnumerateReaderBatches(
                         reader,
                         batchSource,
                         options.GetBatchSize(batchSource.PreferredParallelBatchSize),
                         degreeOfParallelism,
                         batchReader => batchReader.RowsAs(configure),
                         () => EnumerateBatches(
                             reader,
                             genericBatchSize,
                             degreeOfParallelism,
                             values => plan.MapValues(values, culture, dateTimeFormats, typeConverter, errorValuePolicy),
                             captureValues: null,
                             cancellationToken),
                         cancellationToken)) {
                yield return row;
            }
            yield break;
        }

        foreach (T row in EnumerateBatches(
                     reader,
                     genericBatchSize,
                     degreeOfParallelism,
                     values => plan.MapValues(values, culture, dateTimeFormats, typeConverter, errorValuePolicy),
                     captureValues: null,
                     cancellationToken)) {
            yield return row;
        }
    }

    private static IEnumerable<T> EnumerateFactory<T>(
        DbDataReader reader,
        Func<IDataRecord, T> factory,
        ParallelRowMappingOptions? options,
        CancellationToken cancellationToken) {
        options ??= new ParallelRowMappingOptions();
        int degreeOfParallelism = options.GetDegreeOfParallelism();
        if (reader.FieldCount == 0) yield break;

        if (degreeOfParallelism > 1 &&
            reader is IDataReaderParallelBatchSource { CanReadParallelBatches: true } batchSource) {
            foreach (T row in EnumerateFactoryReaderBatches(
                         reader,
                         batchSource,
                         options.GetBatchSize(batchSource.PreferredParallelBatchSize),
                         degreeOfParallelism,
                         factory,
                         cancellationToken)) {
                yield return row;
            }
            yield break;
        }

        foreach (T row in EnumerateSequential(reader.RowsAs(factory), cancellationToken)) {
            yield return row;
        }
    }

    private static IEnumerable<T> EnumerateSequential<T>(
        IEnumerable<T> rows,
        CancellationToken cancellationToken) {
        using IEnumerator<T> enumerator = rows.GetEnumerator();
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!enumerator.MoveNext()) yield break;
            yield return enumerator.Current;
        }
    }

    private static IEnumerable<T> EnumerateBatches<T>(
        DbDataReader reader,
        int batchSize,
        int degreeOfParallelism,
        Func<object?[], T> map,
        Func<DbDataReader, object?[]>? captureValues,
        CancellationToken cancellationToken) {
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var pending = new Queue<Task<T[]>>(degreeOfParallelism);
        ExceptionDispatchInfo? sourceError = null;
        try {
            bool reachedEnd = false;
            while (!reachedEnd) {
                cancellationToken.ThrowIfCancellationRequested();
                RowSnapshotBatch batch;
                try {
                    batch = ReadBatch(
                        reader,
                        batchSize,
                        captureValues,
                        cancellationToken,
                        out reachedEnd);
                } catch (Exception exception) when (!cancellationToken.IsCancellationRequested) {
                    sourceError = ExceptionDispatchInfo.Capture(exception);
                    break;
                }
                if (batch.Count == 0) {
                    batch.Dispose();
                    break;
                }

                pending.Enqueue(Task.Factory.StartNew(
                    () => MapBatch(batch, map, stop.Token),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));

                if (pending.Count < degreeOfParallelism && !reachedEnd) continue;
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
            }

            while (pending.Count > 0) {
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
            }
            sourceError?.Throw();
        } finally {
            stop.Cancel();
            while (pending.Count > 0) {
                Task<T[]> task = pending.Dequeue();
                try {
                    task.GetAwaiter().GetResult();
                } catch {
                    // Observe remaining canceled/faulted work after the ordered exception is propagated.
                }
            }
        }
    }

    private static IEnumerable<T> EnumerateReaderBatches<T>(
        DbDataReader sourceReader,
        IDataReaderParallelBatchSource batchSource,
        int preferredBatchSize,
        int degreeOfParallelism,
        Func<DbDataReader, IEnumerable<T>> mapReader,
        Func<IEnumerable<T>> mapGenericRemainder,
        CancellationToken cancellationToken) {
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
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
                    () => {
                        using (batchReader) {
                            stop.Token.ThrowIfCancellationRequested();
                            return MaterializeRows(mapReader(batchReader), stop.Token);
                        }
                    },
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));

                if (pending.Count < degreeOfParallelism) continue;
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
            }

            while (pending.Count > 0) {
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
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
                    task.GetAwaiter().GetResult();
                } catch {
                    // Observe remaining canceled/faulted work after the ordered exception is propagated.
                }
            }
        }
    }

    private static IEnumerable<T> AwaitNext<T>(
        Queue<Task<T[]>> pending,
        CancellationTokenSource stop,
        CancellationToken cancellationToken) {
        Task<T[]> task = pending.Dequeue();
        T[] rows;
        try {
            rows = task.GetAwaiter().GetResult();
        } catch {
            stop.Cancel();
            throw;
        }
        for (int index = 0; index < rows.Length; index++) {
            if (cancellationToken.CanBeCanceled) {
                cancellationToken.ThrowIfCancellationRequested();
            }
            yield return rows[index];
        }
    }

    private static T[] MaterializeRows<T>(
        IEnumerable<T> rows,
        CancellationToken cancellationToken) {
        var materialized = new List<T>();
        foreach (T row in rows) {
            cancellationToken.ThrowIfCancellationRequested();
            materialized.Add(row);
        }
        return materialized.ToArray();
    }

    private static IEnumerable<T> EnumerateFactoryReaderBatches<T>(
        DbDataReader sourceReader,
        IDataReaderParallelBatchSource batchSource,
        int preferredBatchSize,
        int degreeOfParallelism,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken) {
        using var stop = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
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
                    () => MapFactoryReaderBatch(batchReader, factory, stop.Token),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default));

                if (pending.Count < degreeOfParallelism) continue;
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
            }

            while (pending.Count > 0) {
                foreach (T row in AwaitNext(pending, stop, cancellationToken)) yield return row;
            }

            if (useGenericRemainder) {
                foreach (T row in EnumerateSequential(sourceReader.RowsAs(factory), cancellationToken)) {
                    yield return row;
                }
            }
            sourceError?.Throw();
        } finally {
            stop.Cancel();
            while (pending.Count > 0) {
                Task<T[]> task = pending.Dequeue();
                try {
                    task.GetAwaiter().GetResult();
                } catch {
                    // Observe remaining canceled/faulted work after the ordered exception is propagated.
                }
            }
        }
    }

    private static T[] MapFactoryReaderBatch<T>(
        DbDataReader batchReader,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken) {
        using (batchReader) {
            int expectedCount = batchReader is IDataReaderParallelBatchInfo info
                ? info.ParallelBatchRowCount
                : 0;
            if (expectedCount > 0) {
                var exact = new T[expectedCount];
                int count = 0;
                while (batchReader.Read()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (count == exact.Length) {
                        Array.Resize(ref exact, checked(exact.Length * 2));
                    }
                    exact[count++] = factory(batchReader);
                }
                if (count != exact.Length) Array.Resize(ref exact, count);
                return exact;
            }

            var rows = new List<T>();
            while (batchReader.Read()) {
                cancellationToken.ThrowIfCancellationRequested();
                rows.Add(factory(batchReader));
            }
            return rows.ToArray();
        }
    }

    private static RowSnapshotBatch ReadBatch(
        DbDataReader reader,
        int batchSize,
        Func<DbDataReader, object?[]>? captureValues,
        CancellationToken cancellationToken,
        out bool reachedEnd) {
        var batch = new RowSnapshotBatch(batchSize, reader.FieldCount);
        reachedEnd = false;
        try {
            while (batch.Count < batchSize) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!reader.Read()) {
                    reachedEnd = true;
                    break;
                }
                batch.Add(reader, captureValues);
            }
            return batch;
        } catch {
            batch.Dispose();
            throw;
        }
    }

    private static T[] MapBatch<T>(
        RowSnapshotBatch batch,
        Func<object?[], T> map,
        CancellationToken cancellationToken) {
        try {
            var result = new T[batch.Count];
            for (int index = 0; index < result.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                result[index] = map(batch[index]);
            }
            return result;
        } finally {
            batch.Dispose();
        }
    }

    private sealed class RowSnapshotBatch : IDisposable {
        private readonly object?[][] _rows;
        private readonly bool[] _pooledRows;
        private readonly int _fieldCount;

        internal RowSnapshotBatch(int capacity, int fieldCount) {
            _rows = new object?[capacity][];
            _pooledRows = new bool[capacity];
            _fieldCount = fieldCount;
        }

        internal int Count { get; private set; }
        internal object?[] this[int index] => _rows[index];

        internal void Add(
            DbDataReader reader,
            Func<DbDataReader, object?[]>? captureValues) {
            object?[] values;
            if (captureValues is null) {
#if NET8_0_OR_GREATER
                values = ArrayPool<object?>.Shared.Rent(_fieldCount);
#else
                values = new object?[_fieldCount];
#endif
                int copied = reader.GetValues(values!);
                for (int index = copied; index < _fieldCount; index++) values[index] = DBNull.Value;
            } else {
                values = captureValues(reader);
            }
#if NET8_0_OR_GREATER
            _pooledRows[Count] = captureValues is null;
#else
            _pooledRows[Count] = false;
#endif
            _rows[Count++] = values;
        }

        public void Dispose() {
            for (int index = 0; index < Count; index++) {
                object?[]? values = _rows[index];
                if (values is null) continue;
                _rows[index] = null!;
#if NET8_0_OR_GREATER
                if (_pooledRows[index]) {
                    ArrayPool<object?>.Shared.Return(values, clearArray: true);
                }
#endif
            }
            Count = 0;
        }
    }

}
