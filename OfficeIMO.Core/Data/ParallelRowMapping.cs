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
    internal static bool TryCreateIndependentSnapshotPlan(
        IDataRecord reader,
        out bool[] cloneColumns) {
        return TryCreateIndependentSnapshotPlan(
            reader,
            int.MaxValue,
            out cloneColumns,
            out _);
    }

    internal static bool TryCreateIndependentSnapshotPlan(
        IDataRecord reader,
        int maximumFieldCount,
        out bool[] cloneColumns,
        out bool fieldLimitExceeded) {
        int fieldCount = reader.FieldCount;
        cloneColumns = Array.Empty<bool>();
        fieldLimitExceeded = false;
        if (fieldCount > maximumFieldCount) {
            for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                if (!TryGetIndependentFieldType(reader, ordinal, out _)) return false;
            }
            fieldLimitExceeded = true;
            return true;
        }

        cloneColumns = new bool[fieldCount];
        for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
            if (!TryGetIndependentFieldType(reader, ordinal, out Type fieldType)) return false;
            Type type = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
            cloneColumns[ordinal] = type == typeof(byte[]) || type == typeof(char[]);
        }
        return true;
    }

    private static bool TryGetIndependentFieldType(IDataRecord reader, int ordinal, out Type fieldType) {
        try {
            fieldType = reader.GetFieldType(ordinal);
        } catch (NotSupportedException) {
            fieldType = null!;
            return false;
        } catch (NotImplementedException) {
            fieldType = null!;
            return false;
        }
        return fieldType != null && IsIndependentFieldType(fieldType);
    }

    private static bool IsIndependentFieldType(Type fieldType) {
        Type type = Nullable.GetUnderlyingType(fieldType) ?? fieldType;
        if (type.IsEnum || type.IsPrimitive) return true;
        if (type == typeof(string) || type == typeof(decimal) ||
            type == typeof(DateTime) || type == typeof(DateTimeOffset) ||
            type == typeof(TimeSpan) || type == typeof(Guid) ||
            type == typeof(DBNull) || type == typeof(byte[]) || type == typeof(char[])) {
            return true;
        }
#if NET6_0_OR_GREATER
        if (type == typeof(DateOnly) || type == typeof(TimeOnly)) return true;
#endif
        return false;
    }

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
    /// Readers that expose provider-owned object or other mutable field types also use the sequential
    /// path because those values cannot be copied safely across worker boundaries.
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
    /// <remarks>
    /// Readers that expose provider-owned object or other mutable field types use the sequential
    /// mapping path because those values cannot be copied safely across worker boundaries.
    /// </remarks>
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
    /// A degree of parallelism greater than one snapshots source rows on the calling thread before
    /// invoking the factory on worker tasks. A degree of one preserves the source reader's native
    /// <see cref="IDataRecord"/> behavior on the calling thread. Readers that expose provider-owned
    /// object or other mutable field types also preserve that native sequential behavior.
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

        if (!RowSnapshotSchema.ReaderHasOnlyIndependentFieldTypes(reader)) {
            foreach (T row in EnumerateSequential(reader.RowsAs<T>(), cancellationToken)) {
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

        if (!RowSnapshotSchema.ReaderHasOnlyIndependentFieldTypes(reader)) {
            foreach (T row in EnumerateSequential(reader.RowsAs(configure), cancellationToken)) {
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

        if (degreeOfParallelism == 1) {
            foreach (T row in EnumerateSequential(reader.RowsAs(factory), cancellationToken)) {
                yield return row;
            }
            yield break;
        }

        if (reader is IDataReaderParallelBatchSource { CanReadParallelBatches: true } batchSource) {
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

        foreach (T row in EnumerateFactorySnapshotBatches(
                     reader,
                     options.GetBatchSize(128),
                     degreeOfParallelism,
                     factory,
                     cancellationToken)) {
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

    private static IEnumerable<T> EnumerateFactorySnapshotBatches<T>(
        DbDataReader reader,
        int batchSize,
        int degreeOfParallelism,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken) {
        RowSnapshotSchema? schema = RowSnapshotSchema.TryCapture(reader);
        if (schema is null || !schema.HasOnlyIndependentFieldTypes) {
            foreach (T row in EnumerateSequential(reader.RowsAs(factory), cancellationToken)) {
                yield return row;
            }
            yield break;
        }
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
                        captureValues: null,
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
                    () => MapFactorySnapshotBatch(batch, schema, factory, stop.Token),
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

    private static T[] MapFactorySnapshotBatch<T>(
        RowSnapshotBatch batch,
        RowSnapshotSchema schema,
        Func<IDataRecord, T> factory,
        CancellationToken cancellationToken) {
        try {
            var result = new T[batch.Count];
            var record = new RowSnapshotRecord(schema);
            for (int index = 0; index < result.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                record.SetValues(batch[index]);
                result[index] = factory(record);
            }
            return result;
        } finally {
            batch.Dispose();
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
                for (int index = 0; index < copied; index++) {
                    if (values[index] is byte[] bytes) values[index] = (byte[])bytes.Clone();
                    else if (values[index] is char[] characters) values[index] = (char[])characters.Clone();
                }
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

    private sealed class RowSnapshotSchema {
        private const DynamicallyAccessedMemberTypes CapturedFieldTypeMembers =
            DynamicallyAccessedMemberTypes.PublicFields |
            DynamicallyAccessedMemberTypes.PublicProperties;

        private RowSnapshotSchema(
            string[] names,
            RowSnapshotFieldType[] fieldTypes,
            string[] dataTypeNames,
            CultureInfo culture) {
            Names = names;
            FieldTypes = fieldTypes;
            DataTypeNames = dataTypeNames;
            Culture = culture;
        }

        internal string[] Names { get; }
        internal RowSnapshotFieldType[] FieldTypes { get; }
        internal string[] DataTypeNames { get; }
        internal CultureInfo Culture { get; }
        internal bool HasOnlyIndependentFieldTypes { get; private set; }

        internal static RowSnapshotSchema? TryCapture(DbDataReader reader) {
            int fieldCount = reader.FieldCount;
            var names = new string[fieldCount];
            var fieldTypes = new RowSnapshotFieldType[fieldCount];
            var dataTypeNames = new string[fieldCount];
            try {
                for (int ordinal = 0; ordinal < fieldCount; ordinal++) {
                    names[ordinal] = reader.GetName(ordinal);
                    fieldTypes[ordinal] = new RowSnapshotFieldType(reader.GetFieldType(ordinal));
                    dataTypeNames[ordinal] = reader.GetDataTypeName(ordinal);
                }
            } catch (NotSupportedException) {
                return null;
            } catch (NotImplementedException) {
                return null;
            }
            CultureInfo culture = reader is IDataReaderMappingMetadata metadata
                ? metadata.MappingCulture
                : CultureInfo.InvariantCulture;
            var schema = new RowSnapshotSchema(names, fieldTypes, dataTypeNames, culture) {
                HasOnlyIndependentFieldTypes = fieldTypes.All(
                    fieldType => IsIndependentFieldType(fieldType.Value))
            };
            return schema;
        }

        internal static bool ReaderHasOnlyIndependentFieldTypes(DbDataReader reader) {
            return TryCreateIndependentSnapshotPlan(reader, out _);
        }

        internal sealed class RowSnapshotFieldType {
            internal RowSnapshotFieldType(
                [DynamicallyAccessedMembers(CapturedFieldTypeMembers)] Type value) => Value = value;

            [DynamicallyAccessedMembers(CapturedFieldTypeMembers)]
            internal Type Value { get; }
        }
    }

    private sealed class RowSnapshotRecord : IDataRecord {
        private readonly RowSnapshotSchema _schema;
        private object?[] _values = Array.Empty<object?>();

        internal RowSnapshotRecord(RowSnapshotSchema schema) => _schema = schema;

        internal void SetValues(object?[] values) => _values = values;

        public int FieldCount => _schema.Names.Length;
        public object this[int i] => GetValue(i);
        public object this[string name] => GetValue(GetOrdinal(name));

        public bool GetBoolean(int i) => ConvertValue<bool>(i);
        public byte GetByte(int i) => ConvertValue<byte>(i);
        public char GetChar(int i) => ConvertValue<char>(i);
        public DateTime GetDateTime(int i) => ConvertValue<DateTime>(i);
        public decimal GetDecimal(int i) => ConvertValue<decimal>(i);
        public double GetDouble(int i) => ConvertValue<double>(i);
        public float GetFloat(int i) => ConvertValue<float>(i);
        public Guid GetGuid(int i) {
            object value = GetValue(i);
            return value is Guid guid ? guid : Guid.Parse(Convert.ToString(value, _schema.Culture)!);
        }
        public short GetInt16(int i) => ConvertValue<short>(i);
        public int GetInt32(int i) => ConvertValue<int>(i);
        public long GetInt64(int i) => ConvertValue<long>(i);
        public string GetString(int i) => ConvertValue<string>(i);
        public IDataReader GetData(int i) => throw new NotSupportedException(
            "Nested data readers are not eligible for independent parallel row snapshots.");
        public string GetDataTypeName(int i) => _schema.DataTypeNames[ValidateOrdinal(i)];
        [return: DynamicallyAccessedMembers(
            DynamicallyAccessedMemberTypes.PublicFields |
            DynamicallyAccessedMemberTypes.PublicProperties)]
        public Type GetFieldType(int i) => _schema.FieldTypes[ValidateOrdinal(i)].Value;
        public string GetName(int i) => _schema.Names[ValidateOrdinal(i)];

        public int GetOrdinal(string name) {
            if (name is null) throw new ArgumentNullException(nameof(name));
            for (int index = 0; index < _schema.Names.Length; index++) {
                if (string.Equals(_schema.Names[index], name, StringComparison.Ordinal)) return index;
            }
            for (int index = 0; index < _schema.Names.Length; index++) {
                if (string.Equals(_schema.Names[index], name, StringComparison.OrdinalIgnoreCase)) return index;
            }
            throw new IndexOutOfRangeException($"Column '{name}' was not found.");
        }

        public object GetValue(int i) {
            object? value = _values[ValidateOrdinal(i)];
            return value ?? DBNull.Value;
        }

        public int GetValues(object[] values) {
            if (values is null) throw new ArgumentNullException(nameof(values));
            int count = Math.Min(values.Length, FieldCount);
            for (int index = 0; index < count; index++) values[index] = GetValue(index);
            return count;
        }

        public bool IsDBNull(int i) {
            object? value = _values[ValidateOrdinal(i)];
            return value is null || ReferenceEquals(value, DBNull.Value);
        }

        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferOffset, int length) {
            byte[] source = (byte[])GetValue(i);
            return CopyField(source, fieldOffset, buffer, bufferOffset, length);
        }

        public long GetChars(int i, long fieldOffset, char[]? buffer, int bufferOffset, int length) {
            object value = GetValue(i);
            char[] source = value is string text ? text.ToCharArray() : (char[])value;
            return CopyField(source, fieldOffset, buffer, bufferOffset, length);
        }

        private int ValidateOrdinal(int ordinal) {
            if ((uint)ordinal >= (uint)FieldCount) throw new IndexOutOfRangeException();
            return ordinal;
        }

        private TValue ConvertValue<TValue>(int ordinal) {
            object value = GetValue(ordinal);
            if (value is TValue typed) return typed;
            return (TValue)Convert.ChangeType(value, typeof(TValue), _schema.Culture);
        }

        private static long CopyField<TValue>(
            TValue[] source,
            long fieldOffset,
            TValue[]? buffer,
            int bufferOffset,
            int length) {
            if (fieldOffset < 0 || fieldOffset > source.LongLength) throw new ArgumentOutOfRangeException(nameof(fieldOffset));
            if (buffer is null) return source.LongLength;
            if (bufferOffset < 0 || bufferOffset > buffer.Length) throw new ArgumentOutOfRangeException(nameof(bufferOffset));
            if (length < 0 || length > buffer.Length - bufferOffset) throw new ArgumentOutOfRangeException(nameof(length));
            int available = checked(source.Length - (int)fieldOffset);
            int copied = Math.Min(available, length);
            Array.Copy(source, (int)fieldOffset, buffer, bufferOffset, copied);
            return copied;
        }
    }

}
