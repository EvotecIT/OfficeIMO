using System.Data.Common;
using Apache.Arrow;
using Apache.Arrow.Ipc;
using OfficeIMO.Data;

namespace OfficeIMO.Data.Arrow;

public static partial class DbDataReaderArrowExtensions {
    /// <summary>
    /// Opens the current result set as a bounded Apache Arrow record-batch stream.
    /// </summary>
    /// <remarks>
    /// The stream reads at most <see cref="ArrowReadOptions.BatchSize"/> rows per call.
    /// Disposing the stream does not dispose <paramref name="reader"/>; the caller retains
    /// ownership of the reader and of every returned <see cref="RecordBatch"/>.
    /// Reads are sequential and concurrent calls are rejected.
    /// </remarks>
    public static IArrowArrayStream OpenArrowStream(
        this DbDataReader reader,
        ArrowReadOptions? options = null) {
        ArgumentNullException.ThrowIfNull(reader);
        ArrowReadOptions effectiveOptions = options ?? new ArrowReadOptions();
        Type[]? columnTypes = effectiveOptions.ValidateAndSnapshotColumnTypes(reader.FieldCount);
        ArrowColumnFactory[] columns = CreateColumns(reader, effectiveOptions, columnTypes);
        Schema schema = CreateSchema(reader, columns);
        return new DbDataReaderArrowArrayStream(
            reader,
            columns,
            schema,
            effectiveOptions.BatchSize,
            (reader as IDataReaderFastValueSource)?.FastValueSource);
    }

    /// <summary>
    /// Exports the current result set through the Arrow C stream interface without
    /// materializing the whole result set.
    /// </summary>
    /// <remarks>
    /// Keep the returned owner alive while native code uses its address. Disposing it
    /// releases the exported stream and its unmanaged struct. The source reader remains
    /// caller-owned. Each native <c>get_next</c> call produces at most
    /// <see cref="ArrowReadOptions.BatchSize"/> rows.
    /// </remarks>
    public static ArrowCArrayStreamOwner ExportArrowCStream(
        this DbDataReader reader,
        ArrowReadOptions? options = null) =>
        ArrowCArrayStreamOwner.Export(reader.OpenArrowStream(options));

    private sealed class DbDataReaderArrowArrayStream : IArrowArrayStream {
        private readonly DbDataReader _reader;
        private readonly ArrowColumnFactory[] _columns;
        private readonly int _batchSize;
        private readonly IDataReaderFastValueSource? _fastValueSource;
        private int _readInProgress;
        private bool _completed;
        private bool _disposed;
        private bool _faulted;

        internal DbDataReaderArrowArrayStream(
            DbDataReader reader,
            ArrowColumnFactory[] columns,
            Schema schema,
            int batchSize,
            IDataReaderFastValueSource? fastValueSource) {
            _reader = reader;
            _columns = columns;
            Schema = schema;
            _batchSize = batchSize;
            _fastValueSource = fastValueSource;
        }

        public Schema Schema { get; }

        public async ValueTask<RecordBatch?> ReadNextRecordBatchAsync(
            CancellationToken cancellationToken = default) {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_faulted) {
                throw new InvalidOperationException(
                    "The Arrow stream cannot continue after a failed or cancelled read.");
            }
            cancellationToken.ThrowIfCancellationRequested();
            if (_completed) return null;
            if (Interlocked.CompareExchange(ref _readInProgress, 1, 0) != 0) {
                throw new InvalidOperationException("Arrow stream reads must be sequential.");
            }

            bool readAttempted = false;
            try {
                cancellationToken.ThrowIfCancellationRequested();
                readAttempted = true;
                bool hasRow = await _reader.ReadAsync(cancellationToken).ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
                if (!hasRow) {
                    _completed = true;
                    return null;
                }

                ArrowColumnBuilder[] builders = CreateBuilders(
                    _columns,
                    _batchSize,
                    _fastValueSource);
                int rowCount = 0;
                do {
                    cancellationToken.ThrowIfCancellationRequested();
                    AppendRow(_reader, builders);
                    cancellationToken.ThrowIfCancellationRequested();
                    rowCount++;
                    if (rowCount >= _batchSize) break;

                    readAttempted = true;
                    hasRow = await _reader.ReadAsync(cancellationToken).ConfigureAwait(false);
                    cancellationToken.ThrowIfCancellationRequested();
                } while (hasRow);

                if (!hasRow) _completed = true;
                return BuildBatch(Schema, builders, rowCount, cancellationToken);
            } catch {
                // A forward-only reader may have advanced before the exception. Refuse a
                // retry that could silently omit or duplicate a row.
                if (readAttempted) _faulted = true;
                throw;
            } finally {
                Volatile.Write(ref _readInProgress, 0);
            }
        }

        public void Dispose() {
            _disposed = true;
        }
    }
}
