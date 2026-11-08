using System;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Threading;

namespace OfficeIMO.Data;

public static partial class ParallelRowMappingExtensions {
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
        RowSnapshotPool snapshots,
        Action<DbDataReader, object?[]>? captureValues,
        CancellationToken cancellationToken,
        out bool reachedEnd) {
        RowSnapshotBatch batch = snapshots.Rent();
        reachedEnd = false;
        try {
            while (batch.Count < batch.Capacity) {
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
            if (batch.Columns is IRowMappingSnapshotColumns<T> columns) {
                for (int index = 0; index < result.Length; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    result[index] = columns.Map(index);
                }
            } else {
                for (int index = 0; index < result.Length; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    result[index] = map(batch[index]);
                }
            }
            return result;
        } finally {
            batch.Dispose();
        }
    }

    /// <summary>Owns reusable row storage until the bounded workers have stopped.</summary>
    private sealed class RowSnapshotPool : IDisposable {
        private readonly int _capacity, _fieldCount;
        private readonly Func<int, IRowMappingSnapshotColumns>? _columnsFactory;
        private readonly System.Collections.Concurrent.ConcurrentQueue<RowSnapshotBatch> _available = new();
        // Only the producer creates workspaces. Disposal follows the worker-join finally.
        private readonly List<RowSnapshotBatch> _allocated = new();

        internal RowSnapshotPool(int capacity, int fieldCount, Func<int, IRowMappingSnapshotColumns>? columnsFactory = null) {
            _capacity = capacity;
            _fieldCount = fieldCount;
            _columnsFactory = columnsFactory;
        }

        internal RowSnapshotBatch Rent() {
            if (!_available.TryDequeue(out RowSnapshotBatch? batch)) {
                batch = new RowSnapshotBatch(this, _capacity, _fieldCount, _columnsFactory);
                _allocated.Add(batch);
            }
            batch.Lease();
            return batch;
        }

        internal void Return(RowSnapshotBatch batch) => _available.Enqueue(batch);

        public void Dispose() {
            foreach (RowSnapshotBatch batch in _allocated) batch.ReleaseRows();
            _allocated.Clear();
            while (_available.TryDequeue(out _)) { }
        }
    }

    private sealed class RowSnapshotBatch : IDisposable {
        private readonly RowSnapshotPool _owner;
        private readonly object?[][] _rows;
        private readonly int _fieldCount, _capacity;
        internal IRowMappingSnapshotColumns? Columns { get; }
        private bool _leased;

        internal RowSnapshotBatch(RowSnapshotPool owner, int capacity, int fieldCount,
            Func<int, IRowMappingSnapshotColumns>? columnsFactory) {
            _owner = owner;
            _capacity = capacity;
            Columns = columnsFactory?.Invoke(capacity);
            _rows = Columns is null ? new object?[capacity][] : Array.Empty<object?[]>();
            _fieldCount = fieldCount;
        }

        internal int Count { get; private set; }
        internal int Capacity => _capacity;
        internal object?[] this[int index] => _rows[index];
        internal void Lease() => _leased = true;

        internal void Add(DbDataReader reader, Action<DbDataReader, object?[]>? captureValues) {
            if (Columns is not null) {
                Columns.Capture(reader, Count);
                Count++;
                return;
            }
            object?[]? values = _rows[Count];
            if (values is null) {
#if NET8_0_OR_GREATER
                values = ArrayPool<object?>.Shared.Rent(_fieldCount);
#else
                values = new object?[_fieldCount];
#endif
                _rows[Count] = values;
            }
            if (captureValues is null) {
                int copied = reader.GetValues(values!);
                for (int index = copied; index < _fieldCount; index++) values[index] = DBNull.Value;
            } else {
                for (int index = 0; index < _fieldCount; index++) values[index] = DBNull.Value;
                captureValues(reader, values);
            }
            for (int index = 0; index < _fieldCount; index++) {
                if (values[index] is byte[] bytes) values[index] = (byte[])bytes.Clone();
                else if (values[index] is char[] characters) values[index] = (char[])characters.Clone();
            }
            Count++;
        }

        public void Dispose() {
            if (!_leased) return;
            _leased = false;
            Columns?.Clear();
            // Include an incomplete row whose capture threw before Count advanced.
            foreach (object?[]? values in _rows) {
                if (values is not null) Array.Clear(values, 0, values.Length);
            }
            Count = 0;
            _owner.Return(this);
        }

        internal void ReleaseRows() {
            Columns?.Clear();
            for (int index = 0; index < _rows.Length; index++) {
                object?[]? values = _rows[index];
                if (values is null) continue;
                _rows[index] = null!;
#if NET8_0_OR_GREATER
                ArrayPool<object?>.Shared.Return(values, clearArray: true);
#else
                Array.Clear(values, 0, values.Length);
#endif
            }
            Count = 0;
        }
    }

}
