using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private interface IDirectObjectRows {
            int Count { get; }

            bool HasKnownCount { get; }

            object? GetValue(int rowIndex, int columnIndex);

            void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct);
        }

        private sealed class DirectStreamingObjectRows<T> : IDirectObjectRows {
            private readonly IEnumerable<T> _rows;
            private readonly Action<ExcelTabularRowWriter, T> _writeRow;
            private readonly int _maximumRows;
            private int _count;
            private bool _written;
            private readonly bool _hasKnownCount;

            internal DirectStreamingObjectRows(
                IEnumerable<T> rows,
                Action<ExcelTabularRowWriter, T> writeRow,
                int maximumRows) {
                _rows = rows;
                _writeRow = writeRow;
                _maximumRows = maximumRows;
                if (rows is ICollection<T> collection) {
                    _count = collection.Count;
                    _hasKnownCount = true;
                } else if (rows is IReadOnlyCollection<T> readOnlyCollection) {
                    _count = readOnlyCollection.Count;
                    _hasKnownCount = true;
                }
            }

            public int Count => _count;

            public bool HasKnownCount => _hasKnownCount;

            public object? GetValue(int rowIndex, int columnIndex) =>
                throw new InvalidOperationException(
                    "Single-pass object rows do not support random cell access.");

            public void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct) {
                if (_written) {
                    throw new InvalidOperationException(
                        "Single-pass object rows cannot be written more than once.");
                }
                _written = true;

                bool canCancel = ct.CanBeCanceled;
                int writtenCount = 0;
                using IEnumerator<T> enumerator = _rows.GetEnumerator();
                while (true) {
                    if (canCancel) {
                        ct.ThrowIfCancellationRequested();
                    }
                    if (!enumerator.MoveNext()) {
                        return;
                    }
                    if (writtenCount >= _maximumRows) {
                        throw new InvalidOperationException(
                            "Object-row export exceeds the maximum worksheet row count.");
                    }

                    writer.BeginRow();
                    _writeRow(writer, enumerator.Current);
                    writer.EndRow();
                    writtenCount++;
                    _count = writtenCount;
                }
            }
        }

        private sealed class DirectCallbackRows<T> : IDirectObjectRows {
            private readonly IReadOnlyList<T> _rows;
            private readonly Action<ExcelTabularRowWriter, T> _writeRow;

            internal DirectCallbackRows(IReadOnlyList<T> rows, Action<ExcelTabularRowWriter, T> writeRow) {
                _rows = rows;
                _writeRow = writeRow;
            }

            public int Count => _rows.Count;

            public bool HasKnownCount => true;

            public object? GetValue(int rowIndex, int columnIndex)
                => throw new InvalidOperationException("Streaming callback rows do not support random cell access.");

            public void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < _rows.Count; rowIndex++) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    writer.BeginRow();
                    _writeRow(writer, _rows[rowIndex]);
                    writer.EndRow();
                }
            }
        }
    }
}
