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
            (reader as IDataReaderFastValueSource)?.FastValueSource,
            default);
    }

    /// <summary>
    /// Exports the current result set through the Arrow C stream interface without
    /// materializing the whole result set.
    /// </summary>
    /// <param name="reader">Forward-only reader to project.</param>
    /// <param name="options">Bounded Arrow projection options.</param>
    /// <param name="cancellationToken">Cancellation observed by native stream callbacks.</param>
    /// <remarks>
    /// Acquire a lease from the returned owner for every native call sequence. Disposing
    /// the owner stops new leases and releases the unmanaged struct after active leases
    /// finish. The source reader remains caller-owned. Each native <c>get_next</c> call
    /// produces at most <see cref="ArrowReadOptions.BatchSize"/> rows. Because the Arrow C
    /// stream ABI has no cancellation parameter, <paramref name="cancellationToken"/> is
    /// captured by the exported callbacks.
    /// </remarks>
    public static ArrowCArrayStreamOwner ExportArrowCStream(
        this DbDataReader reader,
        ArrowReadOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(reader);
        ArrowReadOptions effectiveOptions = options ?? new ArrowReadOptions();
        Type[]? columnTypes = effectiveOptions.ValidateAndSnapshotColumnTypes(reader.FieldCount);
        ArrowColumnFactory[] columns = CreateColumns(reader, effectiveOptions, columnTypes);
        Schema schema = CreateSchema(reader, columns);
        return ArrowCArrayStreamOwner.Export(new DbDataReaderArrowArrayStream(
            reader,
            columns,
            schema,
            effectiveOptions.BatchSize,
            (reader as IDataReaderFastValueSource)?.FastValueSource,
            cancellationToken));
    }

    private sealed class DbDataReaderArrowArrayStream : IArrowArrayStream {
        private readonly DbDataReader _reader;
        private readonly ArrowColumnFactory[] _columns;
        private readonly int _batchSize;
        private readonly IDataReaderFastValueSource? _fastValueSource;
        private readonly CancellationToken _ownerCancellationToken;
        private int _readInProgress;
        private bool _completed;
        private bool _disposed;
        private bool _faulted;

        internal DbDataReaderArrowArrayStream(
            DbDataReader reader,
            ArrowColumnFactory[] columns,
            Schema schema,
            int batchSize,
            IDataReaderFastValueSource? fastValueSource,
            CancellationToken ownerCancellationToken) {
            _reader = reader;
            _columns = columns;
            Schema = schema;
            _batchSize = batchSize;
            _fastValueSource = fastValueSource;
            _ownerCancellationToken = ownerCancellationToken;
        }

        public Schema Schema { get; }

        public async ValueTask<RecordBatch?> ReadNextRecordBatchAsync(
            CancellationToken cancellationToken = default) {
            if (Interlocked.CompareExchange(ref _readInProgress, 1, 0) != 0) {
                throw new InvalidOperationException("Arrow stream reads must be sequential.");
            }

            bool readAttempted = false;
            try {
                ObjectDisposedException.ThrowIf(Volatile.Read(ref _disposed), this);
                if (_faulted) {
                    throw new InvalidOperationException(
                        "The Arrow stream cannot continue after a failed or cancelled read.");
                }
                ThrowIfCancellationRequested(cancellationToken);
                if (_completed) return null;
                readAttempted = true;
                CancellationToken effectiveToken = ResolveReadCancellationToken(cancellationToken, out CancellationTokenSource? linkedCancellation);
                using (linkedCancellation) {
                    bool hasRow = await _reader.ReadAsync(effectiveToken).ConfigureAwait(false);
                    ThrowIfCancellationRequested(cancellationToken);
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
                        ThrowIfCancellationRequested(cancellationToken);
                        AppendRow(_reader, builders);
                        ThrowIfCancellationRequested(cancellationToken);
                        rowCount++;
                        if (rowCount >= _batchSize) break;

                        readAttempted = true;
                        hasRow = await _reader.ReadAsync(effectiveToken).ConfigureAwait(false);
                        ThrowIfCancellationRequested(cancellationToken);
                    } while (hasRow);

                    if (!hasRow) _completed = true;
                    return BuildBatch(Schema, builders, rowCount, effectiveToken);
                }
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
            Volatile.Write(ref _disposed, true);
        }

        private void ThrowIfCancellationRequested(CancellationToken readCancellationToken) {
            _ownerCancellationToken.ThrowIfCancellationRequested();
            readCancellationToken.ThrowIfCancellationRequested();
        }

        private CancellationToken ResolveReadCancellationToken(
            CancellationToken readCancellationToken,
            out CancellationTokenSource? linkedCancellation) {
            linkedCancellation = null;
            if (!_ownerCancellationToken.CanBeCanceled) return readCancellationToken;
            if (!readCancellationToken.CanBeCanceled || readCancellationToken == _ownerCancellationToken) {
                return _ownerCancellationToken;
            }

            linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
                _ownerCancellationToken,
                readCancellationToken);
            return linkedCancellation.Token;
        }
    }
}
