using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private interface IDirectObjectRows {
            int Count { get; }

            object? GetValue(int rowIndex, int columnIndex);

            void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct);
        }

        private sealed class DirectObjectRows<T> : IDirectObjectRows {
            private readonly IReadOnlyList<T> _rows;
            private readonly IReadOnlyList<Func<T, object?>> _selectors;

            internal DirectObjectRows(IReadOnlyList<T> rows, IReadOnlyList<Func<T, object?>> selectors) {
                _rows = rows;
                _selectors = selectors;
            }

            public int Count => _rows.Count;

            public object? GetValue(int rowIndex, int columnIndex)
                => _selectors[columnIndex](_rows[rowIndex]);

            public void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < _rows.Count; rowIndex++) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    T row = _rows[rowIndex];
                    writer.BeginRow();
                    for (int columnIndex = 0; columnIndex < _selectors.Count; columnIndex++) {
                        writer.Write(_selectors[columnIndex](row));
                    }
                    writer.EndRow();
                }
            }
        }

        private sealed class DirectTypedObjectRows<T> : IDirectObjectRows {
            private readonly IReadOnlyList<T> _rows;
            private readonly IReadOnlyList<ExcelTabularColumn<T>> _columns;

            internal DirectTypedObjectRows(IReadOnlyList<T> rows, IReadOnlyList<ExcelTabularColumn<T>> columns) {
                _rows = rows;
                _columns = columns;
            }

            public int Count => _rows.Count;

            public object? GetValue(int rowIndex, int columnIndex)
                => _columns[columnIndex].GetValue(_rows[rowIndex]);

            public void WriteRows(ExcelTabularRowWriter writer, CancellationToken ct) {
                bool canCancel = ct.CanBeCanceled;
                for (int rowIndex = 0; rowIndex < _rows.Count; rowIndex++) {
                    if (canCancel) ct.ThrowIfCancellationRequested();
                    T row = _rows[rowIndex];
                    writer.BeginRow();
                    for (int columnIndex = 0; columnIndex < _columns.Count; columnIndex++) {
                        _columns[columnIndex].WriteValue(writer, row);
                    }
                    writer.EndRow();
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
