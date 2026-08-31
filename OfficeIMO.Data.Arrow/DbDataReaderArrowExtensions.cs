using System.Collections.Generic;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using Apache.Arrow;
using Apache.Arrow.Arrays;
using Apache.Arrow.Types;
using OfficeIMO.Data;

namespace OfficeIMO.Data.Arrow;

/// <summary>Apache Arrow projections for forward-only tabular readers.</summary>
public static class DbDataReaderArrowExtensions {
    private const int MaximumInitialReservedCells = 65_536;

    /// <summary>
    /// Converts the current result set into bounded Arrow record batches.
    /// The caller retains ownership of both the reader and every returned batch.
    /// </summary>
    public static IEnumerable<RecordBatch> ReadArrowBatches(
        this DbDataReader reader,
        ArrowReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(reader);
        ArrowReadOptions effectiveOptions = options ?? new ArrowReadOptions();
        Type[]? columnTypes = effectiveOptions.ValidateAndSnapshotColumnTypes(reader.FieldCount);
        ArrowColumnFactory[] columns = CreateColumns(reader, effectiveOptions, columnTypes);
        Schema schema = CreateSchema(reader, columns);
        IDataReaderFastValueSource? fastValueSource =
            (reader as IDataReaderFastValueSource)?.FastValueSource;

        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            bool hasRow = reader.Read();
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasRow) yield break;
            ArrowColumnBuilder[] builders = CreateBuilders(columns, effectiveOptions.BatchSize, fastValueSource);
            int rowCount = 0;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                AppendRow(reader, builders);
                cancellationToken.ThrowIfCancellationRequested();
                rowCount++;
            } while (rowCount < effectiveOptions.BatchSize && reader.Read());

            cancellationToken.ThrowIfCancellationRequested();
            RecordBatch batch = BuildBatch(schema, builders, rowCount, cancellationToken);
            try {
                cancellationToken.ThrowIfCancellationRequested();
            } catch {
                batch.Dispose();
                throw;
            }
            yield return batch;
        }
    }

    /// <summary>
    /// Asynchronously converts the current result set into bounded Arrow record batches.
    /// The caller retains ownership of both the reader and every returned batch.
    /// </summary>
    public static async IAsyncEnumerable<RecordBatch> ReadArrowBatchesAsync(
        this DbDataReader reader,
        ArrowReadOptions? options = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(reader);
        ArrowReadOptions effectiveOptions = options ?? new ArrowReadOptions();
        Type[]? columnTypes = effectiveOptions.ValidateAndSnapshotColumnTypes(reader.FieldCount);
        ArrowColumnFactory[] columns = CreateColumns(reader, effectiveOptions, columnTypes);
        Schema schema = CreateSchema(reader, columns);
        IDataReaderFastValueSource? fastValueSource =
            (reader as IDataReaderFastValueSource)?.FastValueSource;

        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            bool hasRow = await reader.ReadAsync(cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasRow) yield break;
            ArrowColumnBuilder[] builders = CreateBuilders(columns, effectiveOptions.BatchSize, fastValueSource);
            int rowCount = 0;
            do {
                cancellationToken.ThrowIfCancellationRequested();
                AppendRow(reader, builders);
                cancellationToken.ThrowIfCancellationRequested();
                rowCount++;
            } while (rowCount < effectiveOptions.BatchSize &&
                     await reader.ReadAsync(cancellationToken).ConfigureAwait(false));

            cancellationToken.ThrowIfCancellationRequested();
            RecordBatch batch = BuildBatch(schema, builders, rowCount, cancellationToken);
            try {
                cancellationToken.ThrowIfCancellationRequested();
            } catch {
                batch.Dispose();
                throw;
            }
            yield return batch;
        }
    }

    private static ArrowColumnFactory[] CreateColumns(
        DbDataReader reader,
        ArrowReadOptions options,
        IReadOnlyList<Type>? columnTypes) {
        var columns = new ArrowColumnFactory[reader.FieldCount];
        for (int ordinal = 0; ordinal < columns.Length; ordinal++) {
            Type type = columnTypes?[ordinal] ?? reader.GetFieldType(ordinal);
            columns[ordinal] = ArrowColumnFactory.Create(type, options);
        }
        return columns;
    }

    private static Schema CreateSchema(DbDataReader reader, ArrowColumnFactory[] columns) {
        var fields = new Field[columns.Length];
        for (int ordinal = 0; ordinal < fields.Length; ordinal++) {
            fields[ordinal] = new Field(reader.GetName(ordinal), columns[ordinal].ArrowType, nullable: true);
        }
        return new Schema(fields, metadata: null);
    }

    private static ArrowColumnBuilder[] CreateBuilders(
        ArrowColumnFactory[] columns,
        int capacity,
        IDataReaderFastValueSource? fastValueSource) {
        var builders = new ArrowColumnBuilder[columns.Length];
        int initialCapacity = columns.Length == 0
            ? 0
            : Math.Min(capacity, Math.Max(1, MaximumInitialReservedCells / columns.Length));
        for (int ordinal = 0; ordinal < builders.Length; ordinal++) {
            builders[ordinal] = columns[ordinal].CreateBuilder(initialCapacity, fastValueSource);
        }
        return builders;
    }

    private static void AppendRow(DbDataReader reader, ArrowColumnBuilder[] builders) {
        for (int ordinal = 0; ordinal < builders.Length; ordinal++) {
            if (builders[ordinal].HandlesNulls) {
                builders[ordinal].Append(reader, ordinal);
            } else if (reader.IsDBNull(ordinal)) {
                builders[ordinal].AppendNull();
            } else {
                builders[ordinal].Append(reader, ordinal);
            }
        }
    }

    private static RecordBatch BuildBatch(
        Schema schema,
        ArrowColumnBuilder[] builders,
        int rowCount,
        CancellationToken cancellationToken) {
        var arrays = new IArrowArray[builders.Length];
        RecordBatch? batch = null;
        try {
            for (int ordinal = 0; ordinal < arrays.Length; ordinal++) {
                cancellationToken.ThrowIfCancellationRequested();
                arrays[ordinal] = builders[ordinal].Build();
                cancellationToken.ThrowIfCancellationRequested();
            }
            batch = new RecordBatch(schema, arrays, rowCount);
            cancellationToken.ThrowIfCancellationRequested();
            return batch;
        } catch {
            if (batch != null) {
                batch.Dispose();
            } else {
                foreach (IArrowArray? array in arrays) array?.Dispose();
            }
            throw;
        }
    }

    private sealed class ArrowColumnFactory {
        private readonly Func<int, ArrowColumnBuilder> _createBuilder;
        private readonly Func<int, IDataReaderFastValueSource?, ArrowColumnBuilder>? _createFastBuilder;

        private ArrowColumnFactory(IArrowType arrowType, Func<int, ArrowColumnBuilder> createBuilder) {
            ArrowType = arrowType;
            _createBuilder = createBuilder;
        }

        private ArrowColumnFactory(
            IArrowType arrowType,
            Func<int, IDataReaderFastValueSource?, ArrowColumnBuilder> createBuilder) {
            ArrowType = arrowType;
            _createBuilder = capacity => createBuilder(capacity, null);
            _createFastBuilder = createBuilder;
        }

        internal IArrowType ArrowType { get; }

        internal ArrowColumnBuilder CreateBuilder(
            int capacity,
            IDataReaderFastValueSource? fastValueSource) =>
            _createFastBuilder?.Invoke(capacity, fastValueSource) ?? _createBuilder(capacity);

        internal static ArrowColumnFactory Create(Type sourceType, ArrowReadOptions options) {
            Type type = Nullable.GetUnderlyingType(sourceType) ?? sourceType;
            if (type == typeof(bool)) return Primitive<BooleanArray.Builder, BooleanArray, bool>(
                new BooleanType(), static () => new BooleanArray.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetBoolean(i));
            if (type == typeof(sbyte)) return Primitive<Int8Array.Builder, Int8Array, sbyte>(
                new Int8Type(), static () => new Int8Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => Convert.ToSByte(r.GetValue(i), CultureInfo.InvariantCulture));
            if (type == typeof(byte)) return Primitive<UInt8Array.Builder, UInt8Array, byte>(
                new UInt8Type(), static () => new UInt8Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetByte(i));
            if (type == typeof(short)) return Primitive<Int16Array.Builder, Int16Array, short>(
                new Int16Type(), static () => new Int16Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetInt16(i));
            if (type == typeof(ushort)) return Primitive<UInt16Array.Builder, UInt16Array, ushort>(
                new UInt16Type(), static () => new UInt16Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => Convert.ToUInt16(r.GetValue(i), CultureInfo.InvariantCulture));
            if (type == typeof(int)) return Primitive<Int32Array.Builder, Int32Array, int>(
                new Int32Type(), static () => new Int32Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetInt32(i));
            if (type == typeof(uint)) return Primitive<UInt32Array.Builder, UInt32Array, uint>(
                new UInt32Type(), static () => new UInt32Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => Convert.ToUInt32(r.GetValue(i), CultureInfo.InvariantCulture));
            if (type == typeof(long)) {
                return new ArrowColumnFactory(new Int64Type(), (capacity, fastValueSource) => {
                    var builder = new Int64Array.Builder().Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => {
                            if (fastValueSource != null) {
                                if (fastValueSource.TryGetInt64(ordinal, out long value)) {
                                    builder.Append(value);
                                } else {
                                    builder.AppendNull();
                                }
                            } else if (reader.IsDBNull(ordinal)) {
                                builder.AppendNull();
                            } else {
                                builder.Append(reader.GetInt64(ordinal));
                            }
                        },
                        () => builder.Build(),
                        handlesNulls: true);
                });
            }
            if (type == typeof(ulong)) return Primitive<UInt64Array.Builder, UInt64Array, ulong>(
                new UInt64Type(), static () => new UInt64Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => Convert.ToUInt64(r.GetValue(i), CultureInfo.InvariantCulture));
            if (type == typeof(float)) return Primitive<FloatArray.Builder, FloatArray, float>(
                new FloatType(), static () => new FloatArray.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetFloat(i));
            if (type == typeof(double)) {
                return new ArrowColumnFactory(new DoubleType(), (capacity, fastValueSource) => {
                    var builder = new DoubleArray.Builder().Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => {
                            if (fastValueSource != null) {
                                if (fastValueSource.TryGetDouble(ordinal, out double value)) {
                                    builder.Append(value);
                                } else {
                                    builder.AppendNull();
                                }
                            } else if (reader.IsDBNull(ordinal)) {
                                builder.AppendNull();
                            } else {
                                builder.Append(reader.GetDouble(ordinal));
                            }
                        },
                        () => builder.Build(),
                        handlesNulls: true);
                });
            }
            if (type == typeof(decimal)) {
                var decimalType = new Decimal128Type(options.DecimalPrecision, options.DecimalScale);
                return new ArrowColumnFactory(decimalType, capacity => {
                    var builder = new Decimal128Array.Builder(decimalType).Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => builder.Append(reader.GetDecimal(ordinal)),
                        () => builder.Build());
                });
            }
            if (type == typeof(DateTime)) {
                var timestampType = new TimestampType(TimeUnit.Microsecond, (string)null!);
                return new ArrowColumnFactory(timestampType, (capacity, fastValueSource) => {
                    var builder = new TimestampArray.Builder(timestampType).Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => {
                            if (fastValueSource != null) {
                                if (fastValueSource.TryGetDateTime(ordinal, out DateTime value)) {
                                    builder.Append(ToTimezoneLessTimestamp(value));
                                } else {
                                    builder.AppendNull();
                                }
                            } else if (reader.IsDBNull(ordinal)) {
                                builder.AppendNull();
                            } else {
                                builder.Append(ToTimezoneLessTimestamp(reader.GetDateTime(ordinal)));
                            }
                        },
                        () => builder.Build(),
                        handlesNulls: true);
                });
            }
            if (type == typeof(DateTimeOffset)) {
                var timestampType = new TimestampType(TimeUnit.Microsecond, TimeZoneInfo.Utc);
                return new ArrowColumnFactory(timestampType, capacity => {
                    var builder = new TimestampArray.Builder(timestampType).Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => builder.Append(reader.GetFieldValue<DateTimeOffset>(ordinal).ToUniversalTime()),
                        () => builder.Build());
                });
            }
            if (type == typeof(DateOnly)) return Primitive<Date32Array.Builder, Date32Array, DateOnly>(
                new Date32Type(), static () => new Date32Array.Builder(), static (b, c) => b.Reserve(c), static b => b.AppendNull(), static b => b.Build(), static (b, v) => b.Append(v), static (r, i) => r.GetFieldValue<DateOnly>(i));
            if (type == typeof(TimeOnly)) {
                var timeType = new Time64Type(TimeUnit.Microsecond);
                return new ArrowColumnFactory(timeType, capacity => {
                    var builder = new Time64Array.Builder(timeType).Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => builder.Append(reader.GetFieldValue<TimeOnly>(ordinal)),
                        () => builder.Build());
                });
            }
            if (type == typeof(Guid)) {
                var guidType = new FixedSizeBinaryType(16);
                return new ArrowColumnFactory(guidType, capacity => {
                    var builder = new GuidArray.Builder().Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => builder.Append(reader.GetGuid(ordinal)),
                        () => builder.Build());
                });
            }
            if (type == typeof(byte[])) {
                return new ArrowColumnFactory(new BinaryType(), capacity => {
                    var builder = new BinaryArray.Builder().Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => builder.Append((byte[])reader.GetValue(ordinal)),
                        () => builder.Build());
                });
            }
            if (type == typeof(string) || options.ConvertUnsupportedTypesToString) {
                return new ArrowColumnFactory(new StringType(), (capacity, fastValueSource) => {
                    var builder = new StringArray.Builder().Reserve(capacity);
                    return new ArrowColumnBuilder(
                        () => builder.AppendNull(),
                        (reader, ordinal) => {
                            if (type == typeof(string) && fastValueSource != null) {
                                if (fastValueSource.TryGetUtf8Value(ordinal, out ArraySegment<byte> value)) {
                                    builder.Append(value.Array!.AsSpan(value.Offset, value.Count));
                                } else if (reader.IsDBNull(ordinal)) {
                                    builder.AppendNull();
                                } else {
                                    builder.Append(reader.GetString(ordinal), Encoding.UTF8);
                                }
                            } else if (reader.IsDBNull(ordinal)) {
                                builder.AppendNull();
                            } else {
                                builder.Append(
                                    Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty,
                                    Encoding.UTF8);
                            }
                        },
                        () => builder.Build(),
                        handlesNulls: true);
                });
            }

            throw new NotSupportedException($"CLR type '{type.FullName}' does not have an Apache Arrow adapter.");
        }

        private static ArrowColumnFactory Primitive<TBuilder, TArray, TValue>(
            IArrowType arrowType,
            Func<TBuilder> create,
            Func<TBuilder, int, TBuilder> reserve,
            Func<TBuilder, TBuilder> appendNull,
            Func<TBuilder, TArray> build,
            Action<TBuilder, TValue> append,
            Func<DbDataReader, int, TValue> read)
            where TArray : IArrowArray {
            return new ArrowColumnFactory(arrowType, capacity => {
                TBuilder builder = reserve(create(), capacity);
                return new ArrowColumnBuilder(
                    () => appendNull(builder),
                    (reader, ordinal) => append(builder, read(reader, ordinal)),
                    () => build(builder));
            });
        }

        private static DateTimeOffset ToTimezoneLessTimestamp(DateTime value) =>
            new(DateTime.SpecifyKind(value, DateTimeKind.Utc));
    }

    private sealed class ArrowColumnBuilder {
        private readonly Action _appendNull;
        private readonly Action<DbDataReader, int> _append;
        private readonly Func<IArrowArray> _build;

        internal ArrowColumnBuilder(
            Action appendNull,
            Action<DbDataReader, int> append,
            Func<IArrowArray> build,
            bool handlesNulls = false) {
            _appendNull = appendNull;
            _append = append;
            _build = build;
            HandlesNulls = handlesNulls;
        }

        internal bool HandlesNulls { get; }

        internal void AppendNull() => _appendNull();

        internal void Append(DbDataReader reader, int ordinal) => _append(reader, ordinal);

        internal IArrowArray Build() => _build();
    }
}
