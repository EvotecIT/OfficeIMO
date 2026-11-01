namespace OfficeIMO.Excel {
    internal enum ExcelBufferedCellKind : byte {
        Blank,
        Text,
        Number,
        Int32,
        Int64,
        DateTime,
        Boolean,
        Value
    }

    /// <summary>A cell value carried without boxing the common primitive types.</summary>
    internal readonly struct ExcelBufferedCell {
        private const long TicksMask = 0x3FFF_FFFF_FFFF_FFFFL;
        private const int KindShift = 62;

        private ExcelBufferedCell(ExcelBufferedCellKind kind, long bits, object? reference) {
            Kind = kind;
            Bits = bits;
            Reference = reference;
        }

        internal ExcelBufferedCellKind Kind { get; }

        internal long Bits { get; }

        internal object? Reference { get; }

        internal double NumberValue => BitConverter.Int64BitsToDouble(Bits);

        internal DateTime DateTimeValue => new DateTime(Bits & TicksMask, (DateTimeKind)((ulong)Bits >> KindShift));

        internal static ExcelBufferedCell Text(string? value) => value == null
            ? default
            : new ExcelBufferedCell(ExcelBufferedCellKind.Text, 0L, value);

        internal static ExcelBufferedCell Number(double value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Number, BitConverter.DoubleToInt64Bits(value), null);

        internal static ExcelBufferedCell Int32(int value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Int32, value, null);

        internal static ExcelBufferedCell Int64(long value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Int64, value, null);

        internal static ExcelBufferedCell Boolean(bool value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Boolean, value ? 1L : 0L, null);

        internal static ExcelBufferedCell DateTimeCell(DateTime value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.DateTime, value.Ticks | ((long)value.Kind << KindShift), null);

        internal static ExcelBufferedCell FromValue(object? value) {
            switch (value) {
                case null:
                case DBNull:
                    return default;
                case string text:
                    return Text(text);
                case double number:
                    return Number(number);
                case int number:
                    return Int32(number);
                case long number:
                    return Int64(number);
                case bool boolean:
                    return Boolean(boolean);
                case DateTime dateTime:
                    return DateTimeCell(dateTime);
                default:
                    return new ExcelBufferedCell(ExcelBufferedCellKind.Value, 0L, value);
            }
        }

        internal static ExcelBufferedCell Restore(ExcelBufferedCellKind kind, long bits, object? reference)
            => new ExcelBufferedCell(kind, bits, reference);

        internal object? ToValue() {
            switch (Kind) {
                case ExcelBufferedCellKind.Text:
                case ExcelBufferedCellKind.Value:
                    return Reference;
                case ExcelBufferedCellKind.Number:
                    return NumberValue;
                case ExcelBufferedCellKind.Int32:
                    return (int)Bits;
                case ExcelBufferedCellKind.Int64:
                    return Bits;
                case ExcelBufferedCellKind.DateTime:
                    return DateTimeValue;
                case ExcelBufferedCellKind.Boolean:
                    return Bits != 0L;
                default:
                    return null;
            }
        }
    }

    /// <summary>
    /// Holds cells written in ascending row and column order for a worksheet that has not
    /// reached the DOM. Storage grows in fixed chunks, so nothing is copied as rows accumulate
    /// and a cell costs twelve bytes plus its text reference.
    /// </summary>
    internal sealed class ExcelCellValueBuffer {
        private const int RectangleCellRatioLimit = 8;
        private const int MinimumRectangleCellLimit = 1 << 20;
        private const int KindShift = 16;
        private const int ColumnMask = 0xFFFF;
        private const byte MixedColumn = byte.MaxValue;

        private readonly struct RowEntry {
            internal RowEntry(int rowNumber, int firstCell) {
                RowNumber = rowNumber;
                FirstCell = firstCell;
            }

            internal int RowNumber { get; }

            internal int FirstCell { get; }
        }

        private sealed class ChunkedList<T> {
            private const int ChunkBits = 12;
            private const int ChunkSize = 1 << ChunkBits;
            private const int ChunkMask = ChunkSize - 1;
            private T[][] _chunks = new T[8][];

            internal int Count { get; private set; }

            internal T this[int index] => _chunks[index >> ChunkBits][index & ChunkMask];

            internal void Add(T value) {
                int chunkIndex = Count >> ChunkBits;
                if ((Count & ChunkMask) == 0) {
                    if (chunkIndex == _chunks.Length) {
                        Array.Resize(ref _chunks, _chunks.Length * 2);
                    }

                    _chunks[chunkIndex] = new T[ChunkSize];
                }

                _chunks[chunkIndex][Count & ChunkMask] = value;
                Count++;
            }
        }

        private readonly ChunkedList<long> _payloads = new ChunkedList<long>();
        private readonly ChunkedList<int> _headers = new ChunkedList<int>();
        private readonly ChunkedList<object> _references = new ChunkedList<object>();
        private readonly ChunkedList<RowEntry> _rows = new ChunkedList<RowEntry>();
        private byte[] _columnShapes = new byte[16];
        private int _firstColumn = int.MaxValue;
        private int _lastColumn;
        private int _maximumColumn;

        /// <summary>Gets the last written row; the tabular view of the buffer starts at row one.</summary>
        internal int RowCount { get; private set; }

        /// <summary>Gets the last written column; the tabular view of the buffer starts at column one.</summary>
        internal int ColumnCount => _maximumColumn;

        internal int CellCount => _headers.Count;

        internal int WrittenRowCount => _rows.Count;

        /// <summary>
        /// Gets whether the cells can stand as a worksheet. A buffer far sparser than its
        /// rectangle stays out, because format writers walk the rectangle cell by cell.
        /// </summary>
        internal bool HasCompactRectangle
            => (long)RowCount * _maximumColumn <= Math.Max(MinimumRectangleCellLimit, (long)CellCount * RectangleCellRatioLimit);

        internal string Dimension
            => ExcelA1.CellReference(_rows[0].RowNumber, _firstColumn) + ":" + ExcelA1.CellReference(RowCount, _maximumColumn);

        internal bool CanAdd(int row, int column) {
            if (row <= 0 || row > ExcelA1.MaxRows || column <= 0 || column > ExcelA1.MaxColumns) {
                return false;
            }

            return row > RowCount || (row == RowCount && column > _lastColumn);
        }

        internal void Add(int row, int column, in ExcelBufferedCell cell) {
            if (row != RowCount) {
                _rows.Add(new RowEntry(row, _headers.Count));
                RowCount = row;
            }

            long payload = cell.Bits;
            if (cell.Reference != null) {
                payload = _references.Count;
                _references.Add(cell.Reference);
            }

            _payloads.Add(payload);
            _headers.Add(((int)cell.Kind << KindShift) | column);
            _lastColumn = column;
            if (column > _maximumColumn) {
                _maximumColumn = column;
            }

            if (column < _firstColumn) {
                _firstColumn = column;
            }

            NoteColumnShape(column, cell.Kind);
        }

        private void NoteColumnShape(int column, ExcelBufferedCellKind kind) {
            if (kind == ExcelBufferedCellKind.Blank) {
                return;
            }

            if (column > _columnShapes.Length) {
                Array.Resize(ref _columnShapes, Math.Min(ExcelA1.MaxColumns, Math.Max(column, _columnShapes.Length * 2)));
            }

            byte shape = (byte)kind;
            ref byte current = ref _columnShapes[column - 1];
            if (current == 0) {
                current = shape;
            } else if (current != shape) {
                current = MixedColumn;
            }
        }

        internal Type[] CreateColumnTypes() {
            var types = new Type[_maximumColumn];
            for (int i = 0; i < types.Length; i++) {
                types[i] = (i < _columnShapes.Length ? _columnShapes[i] : (byte)0) switch {
                    (byte)ExcelBufferedCellKind.Text => typeof(string),
                    (byte)ExcelBufferedCellKind.Number => typeof(double),
                    (byte)ExcelBufferedCellKind.Int32 => typeof(int),
                    (byte)ExcelBufferedCellKind.Int64 => typeof(long),
                    (byte)ExcelBufferedCellKind.DateTime => typeof(DateTime),
                    (byte)ExcelBufferedCellKind.Boolean => typeof(bool),
                    _ => typeof(object)
                };
            }

            return types;
        }

        internal int GetRowNumber(int writtenRowIndex) => _rows[writtenRowIndex].RowNumber;

        internal int GetFirstCellIndex(int writtenRowIndex) => _rows[writtenRowIndex].FirstCell;

        internal int GetEndCellIndex(int writtenRowIndex)
            => writtenRowIndex + 1 < _rows.Count ? _rows[writtenRowIndex + 1].FirstCell : _headers.Count;

        internal int GetColumn(int cellIndex) => _headers[cellIndex] & ColumnMask;

        internal ExcelBufferedCell GetCell(int cellIndex) {
            var kind = (ExcelBufferedCellKind)(_headers[cellIndex] >> KindShift);
            long payload = _payloads[cellIndex];
            return kind == ExcelBufferedCellKind.Text || kind == ExcelBufferedCellKind.Value
                ? ExcelBufferedCell.Restore(kind, 0L, _references[(int)payload])
                : ExcelBufferedCell.Restore(kind, payload, null);
        }

        /// <summary>
        /// Reads the tabular view. An omitted cell is null; an explicit null keeps the
        /// CellValue contract of an empty text cell.
        /// </summary>
        internal object? GetValue(int rowIndex, int columnIndex) {
            int cellIndex = FindCell(rowIndex + 1, columnIndex + 1);
            if (cellIndex < 0) {
                return null;
            }

            ExcelBufferedCell cell = GetCell(cellIndex);
            return cell.Kind == ExcelBufferedCellKind.Blank ? string.Empty : cell.ToValue();
        }

        internal bool Contains(int row, int column) => FindCell(row, column) >= 0;

        private int FindCell(int row, int column) {
            int writtenRow = FindWrittenRow(row);
            if (writtenRow < 0) {
                return -1;
            }

            int low = GetFirstCellIndex(writtenRow);
            int high = GetEndCellIndex(writtenRow) - 1;
            while (low <= high) {
                int middle = low + ((high - low) >> 1);
                int candidate = GetColumn(middle);
                if (candidate == column) {
                    return middle;
                }

                if (candidate < column) {
                    low = middle + 1;
                } else {
                    high = middle - 1;
                }
            }

            return -1;
        }

        private int FindWrittenRow(int rowNumber) {
            if (_rows.Count == RowCount) {
                return rowNumber <= RowCount ? rowNumber - 1 : -1;
            }

            int low = 0;
            int high = _rows.Count - 1;
            while (low <= high) {
                int middle = low + ((high - low) >> 1);
                int candidate = _rows[middle].RowNumber;
                if (candidate == rowNumber) {
                    return middle;
                }

                if (candidate < rowNumber) {
                    low = middle + 1;
                } else {
                    high = middle - 1;
                }
            }

            return -1;
        }

        internal IEnumerable<(int Row, int Column, object? Value)> EnumerateWrittenCells() {
            for (int rowIndex = 0; rowIndex < _rows.Count; rowIndex++) {
                int rowNumber = GetRowNumber(rowIndex);
                int end = GetEndCellIndex(rowIndex);
                for (int cellIndex = GetFirstCellIndex(rowIndex); cellIndex < end; cellIndex++) {
                    yield return (rowNumber, GetColumn(cellIndex), GetCell(cellIndex).ToValue());
                }
            }
        }

        /// <summary>Gets the text values in write order.</summary>
        internal IEnumerable<string> EnumerateText() {
            for (int i = 0; i < _references.Count; i++) {
                if (_references[i] is string text) {
                    yield return text;
                }
            }
        }
    }
}
