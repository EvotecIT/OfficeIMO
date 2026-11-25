namespace OfficeIMO.Excel {
    internal enum ExcelBufferedCellKind : byte {
        Blank,
        Text,
        Number,
        Int32,
        Int64,
        DateTime,
        Boolean,
        Value,
        StyleOnly
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

        internal static ExcelBufferedCell StyleOnly { get; } = new ExcelBufferedCell(ExcelBufferedCellKind.StyleOnly, 0L, null);

        internal static ExcelBufferedCell Number(double value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Number, BitConverter.DoubleToInt64Bits(value), null);

        internal static ExcelBufferedCell Int32(int value)
            => new ExcelBufferedCell(ExcelBufferedCellKind.Int32, value, null);

        internal static ExcelBufferedCell Int64(long value)
            => ExcelExactNumber.FitsNumberCell(value)
                ? new ExcelBufferedCell(ExcelBufferedCellKind.Int64, value, null)
                : Text(InvariantNumberText.Get(value));

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

        internal bool HasLineBreak => Reference is string text && (text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0);

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
    /// <remarks>
    /// The last rows can instead be a block of typed selector values (<see cref="AddBlock"/>), kept as the caller's
    /// <see cref="ExcelTabularCellValues"/> rather than copied cell by cell. Cell indexes run on through the block, so
    /// readers and writers see the same cells either way. The first write that changes a block cell, or adds a cell
    /// after it, copies the block into the chunks.
    /// </remarks>
    internal sealed class ExcelCellValueBuffer {
        private const int RectangleCellRatioLimit = 8;
        private const int MinimumRectangleCellLimit = 1 << 20;
        private const int KindShift = 16;
        private const int ColumnMask = 0xFFFF;
        private const byte MixedColumn = byte.MaxValue;
        private const int StyleChunkBits = 12;
        private const int StyleChunkMask = (1 << StyleChunkBits) - 1;
        private const int ReadCopyRatioLimit = 8;
        private const int MinimumReadCopyLimit = 1 << 16;

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

            internal T this[int index] {
                get => _chunks[index >> ChunkBits][index & ChunkMask];
                set => _chunks[index >> ChunkBits][index & ChunkMask] = value;
            }

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

        private uint[]?[] _styleChunks = new uint[]?[8];
        private readonly ChunkedList<long> _payloads = new ChunkedList<long>();
        private readonly ChunkedList<int> _headers = new ChunkedList<int>();
        private readonly ChunkedList<object> _references = new ChunkedList<object>();
        private readonly ChunkedList<RowEntry> _rows = new ChunkedList<RowEntry>();
        private byte[] _columnShapes = new byte[16];
        private int _firstColumn = int.MaxValue;
        private int _lastColumn;
        private int _maximumColumn;
        private long _readCopiedCells;
        // The trailing block of typed values: whole rows from column one, none blank and none holding an arbitrary object.
        private ExcelTabularCellValues? _block;
        private int _blockFirstValue;
        private int _blockFirstRow;
        private int _blockRows;
        private int _blockColumns;
        private uint _blockDateStyle;
        private int _rowCountBeforeBlock;
        private int _lastColumnBeforeBlock;

        /// <summary>Gets a number that changes whenever the buffer's content does.</summary>
        internal int Version { get; private set; }

        /// <summary>Gets the last written row; the tabular view of the buffer starts at row one.</summary>
        internal int RowCount { get; private set; }

        /// <summary>Gets the last written column; the tabular view of the buffer starts at column one.</summary>
        internal int ColumnCount => _maximumColumn;

        internal int CellCount => _headers.Count + _blockRows * _blockColumns;

        internal int WrittenRowCount => _rows.Count + _blockRows;

        /// <summary>The number of cells before the block; every cell when there is none.</summary>
        private int ChunkedCellCount => _headers.Count;

        private bool IsBlockCell(int cellIndex) => _block != null && cellIndex >= _headers.Count;

        /// <summary>
        /// Gets whether the cells can stand as a worksheet. A buffer far sparser than its
        /// rectangle stays out, because format writers walk the rectangle cell by cell.
        /// </summary>
        internal bool HasCompactRectangle
            => (long)RowCount * _maximumColumn <= Math.Max(MinimumRectangleCellLimit, (long)CellCount * RectangleCellRatioLimit);

        internal string TabularRange => "A1:" + ExcelA1.CellReference(RowCount, _maximumColumn);

        internal string Dimension {
            get {
                string first = ExcelA1.CellReference(GetRowNumber(0), _firstColumn);
                string last = ExcelA1.CellReference(RowCount, _maximumColumn);
                return first == last ? first : first + ":" + last;
            }
        }

        /// <summary>
        /// Admits one more copy of the buffer for a reader. A loop that reads after every write would
        /// otherwise cost quadratic time, so such a loop runs out and moves the cells to the DOM.
        /// </summary>
        internal bool TryReserveReadCopy() {
            long copied = _readCopiedCells + CellCount;
            if (copied > Math.Max(MinimumReadCopyLimit, (long)CellCount * ReadCopyRatioLimit)) {
                return false;
            }

            _readCopiedCells = copied;
            return true;
        }

        internal bool CanAdd(int row, int column) {
            if (row <= 0 || row > ExcelA1.MaxRows || column <= 0 || column > ExcelA1.MaxColumns) {
                return false;
            }

            return row > RowCount || (row == RowCount && column > _lastColumn);
        }

        internal void Add(int row, int column, in ExcelBufferedCell cell) {
            ExpandBlock();
            Version++;
            if (row != RowCount) {
                _rows.Add(new RowEntry(row, _headers.Count));
                RowCount = row;
            }

            _payloads.Add(StorePayload(cell));
            _headers.Add(CreateHeader(column, cell));
            _lastColumn = column;
            if (column > _maximumColumn) {
                _maximumColumn = column;
            }

            if (column < _firstColumn) {
                _firstColumn = column;
            }

            NoteColumnShape(column, cell.Kind);
        }

        /// <summary>Replaces an existing cell, keeping its position.</summary>
        internal void Replace(int cellIndex, in ExcelBufferedCell cell) {
            if (IsBlockCell(cellIndex)) ExpandBlock();
            Version++;
            int column = GetColumn(cellIndex);
            _payloads[cellIndex] = StorePayload(cell);
            _headers[cellIndex] = CreateHeader(column, cell);
            NoteColumnShape(column, cell.Kind);
        }

        private long StorePayload(in ExcelBufferedCell cell) {
            if (cell.Reference == null) {
                return cell.Bits;
            }

            _references.Add(cell.Reference);
            return _references.Count - 1;
        }

        private static int CreateHeader(int column, in ExcelBufferedCell cell) => ((int)cell.Kind << KindShift) | column;

        /// <summary>Gets whether any cell carries a workbook cell format.</summary>
        internal bool HasCellStyles { get; private set; }

        /// <summary>Gets the workbook cell format of a cell; zero leaves it to its value.</summary>
        internal uint GetStyle(int cellIndex) {
            if (IsBlockCell(cellIndex)) {
                return GetBlockKind(cellIndex) == ExcelBufferedCellKind.DateTime ? _blockDateStyle : 0U;
            }

            int chunkIndex = cellIndex >> StyleChunkBits;
            uint[]? chunk = chunkIndex < _styleChunks.Length ? _styleChunks[chunkIndex] : null;
            return chunk == null ? 0U : chunk[cellIndex & StyleChunkMask];
        }

        // Styles sit beside the cells in chunks made on first use, so a large unstyled range
        // costs nothing and an index can reach every format a workbook may hold.
        internal void SetStyle(int cellIndex, uint styleIndex) {
            if (IsBlockCell(cellIndex)) ExpandBlock();
            Version++;
            int chunkIndex = cellIndex >> StyleChunkBits;
            if (chunkIndex >= _styleChunks.Length) {
                Array.Resize(ref _styleChunks, Math.Max(chunkIndex + 1, _styleChunks.Length * 2));
            }

            uint[]? chunk = _styleChunks[chunkIndex];
            if (chunk == null) {
                if (styleIndex == 0U) {
                    return;
                }

                chunk = _styleChunks[chunkIndex] = new uint[1 << StyleChunkBits];
            }

            chunk[cellIndex & StyleChunkMask] = styleIndex;
            HasCellStyles |= styleIndex != 0U;
        }

        private void NoteColumnShape(int column, ExcelBufferedCellKind kind) {
            if (kind == ExcelBufferedCellKind.Blank || kind == ExcelBufferedCellKind.StyleOnly) {
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

        internal int GetRowNumber(int writtenRowIndex)
            => writtenRowIndex < _rows.Count ? _rows[writtenRowIndex].RowNumber : _blockFirstRow + (writtenRowIndex - _rows.Count);

        internal int GetFirstCellIndex(int writtenRowIndex)
            => writtenRowIndex < _rows.Count
                ? _rows[writtenRowIndex].FirstCell
                : ChunkedCellCount + (writtenRowIndex - _rows.Count) * _blockColumns;

        internal int GetEndCellIndex(int writtenRowIndex)
            => writtenRowIndex + 1 < _rows.Count
                ? _rows[writtenRowIndex + 1].FirstCell
                : writtenRowIndex < _rows.Count
                    ? ChunkedCellCount
                    : ChunkedCellCount + (writtenRowIndex - _rows.Count + 1) * _blockColumns;

        internal int GetColumn(int cellIndex)
            => IsBlockCell(cellIndex) ? (cellIndex - ChunkedCellCount) % _blockColumns + 1 : _headers[cellIndex] & ColumnMask;

        internal ExcelBufferedCell GetCell(int cellIndex) {
            if (IsBlockCell(cellIndex)) {
                int offset = _block!.GetChunk(_blockFirstValue + cellIndex - ChunkedCellCount, out byte[] kinds, out long[] bits);
                var blockKind = (ExcelBufferedCellKind)kinds[offset];
                return blockKind == ExcelBufferedCellKind.Text
                    ? ExcelBufferedCell.Restore(blockKind, 0L, _block.GetReferenceById(bits[offset]))
                    : ExcelBufferedCell.Restore(blockKind, bits[offset], null);
            }

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

        internal int FindCell(int row, int column) {
            if (_block != null && row >= _blockFirstRow) {
                return row < _blockFirstRow + _blockRows && column >= 1 && column <= _blockColumns
                    ? ChunkedCellCount + (row - _blockFirstRow) * _blockColumns + column - 1
                    : -1;
            }

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
            int chunkedRowCount = _block != null ? _rowCountBeforeBlock : RowCount;
            if (_rows.Count == chunkedRowCount) {
                return rowNumber <= chunkedRowCount ? rowNumber - 1 : -1;
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
            for (int rowIndex = 0; rowIndex < WrittenRowCount; rowIndex++) {
                int rowNumber = GetRowNumber(rowIndex);
                int end = GetEndCellIndex(rowIndex);
                for (int cellIndex = GetFirstCellIndex(rowIndex); cellIndex < end; cellIndex++) {
                    yield return (rowNumber, GetColumn(cellIndex), GetCell(cellIndex).ToValue());
                }
            }
        }

        /// <summary>Gets the text values in write order.</summary>
        internal IEnumerable<string> EnumerateText() {
            for (int rowIndex = 0; rowIndex < WrittenRowCount; rowIndex++) {
                int end = GetEndCellIndex(rowIndex);
                for (int cellIndex = GetFirstCellIndex(rowIndex); cellIndex < end; cellIndex++) {
                    if (GetCell(cellIndex).Reference is string text) {
                        yield return text;
                    }
                }
            }
        }

        /// <summary>
        /// Appends <paramref name="rowCount"/> whole rows of <paramref name="values"/>, from value
        /// <paramref name="firstValue"/>, as rows <paramref name="firstRow"/> onwards, keeping the values where they are.
        /// Date cells take <paramref name="dateStyle"/>. Returns false, with nothing added, when the rows do not follow
        /// the buffer, a cell is blank or holds an arbitrary object, or text carries a line break, which takes an
        /// automatic style of its own.
        /// </summary>
        internal bool AddBlock(int firstRow, ExcelTabularCellValues values, int firstValue, int rowCount, uint dateStyle) {
            int columns = values.ColumnCount;
            if (_block != null || rowCount <= 0 || firstRow <= RowCount || (long)firstRow + rowCount - 1 > ExcelA1.MaxRows
                || firstValue < 0 || (long)firstValue + (long)rowCount * columns > (long)values.RowCount * columns) {
                return false;
            }

            int end = firstValue + rowCount * columns;
            bool hasDates = false;
            for (int index = firstValue; index < end; index++) {
                switch (values.GetKind(index)) {
                    case ExcelBufferedCellKind.Text:
                        string text = (string)values.GetReference(index);
                        if (text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0) return false;
                        break;
                    case ExcelBufferedCellKind.DateTime:
                        hasDates = true;
                        break;
                    case ExcelBufferedCellKind.Number:
                    case ExcelBufferedCellKind.Int32:
                    case ExcelBufferedCellKind.Int64:
                    case ExcelBufferedCellKind.Boolean:
                        break;
                    default:
                        return false;
                }
            }

            Version++;
            _block = values;
            _blockFirstValue = firstValue;
            _blockFirstRow = firstRow;
            _blockRows = rowCount;
            _blockColumns = columns;
            _blockDateStyle = dateStyle;
            _rowCountBeforeBlock = RowCount;
            _lastColumnBeforeBlock = _lastColumn;
            RowCount = firstRow + rowCount - 1;
            _lastColumn = columns;
            if (columns > _maximumColumn) _maximumColumn = columns;
            _firstColumn = 1;
            HasCellStyles |= hasDates && dateStyle != 0U;
            for (int index = firstValue, column = 1; index < end; index++) {
                NoteColumnShape(column, values.GetKind(index));
                column = column == columns ? 1 : column + 1;
            }

            return true;
        }

        private ExcelBufferedCellKind GetBlockKind(int cellIndex)
            => _block!.GetKind(_blockFirstValue + cellIndex - ChunkedCellCount);

        /// <summary>Copies the block into the chunks, as if its cells had been added one by one.</summary>
        private void ExpandBlock() {
            ExcelTabularCellValues? values = _block;
            if (values == null) {
                return;
            }

            int firstValue = _blockFirstValue, firstRow = _blockFirstRow, rows = _blockRows, columns = _blockColumns;
            uint dateStyle = _blockDateStyle;
            _block = null;
            _blockRows = 0;
            _blockColumns = 0;
            RowCount = _rowCountBeforeBlock;
            _lastColumn = _lastColumnBeforeBlock;
            int index = firstValue;
            for (int row = firstRow; row < firstRow + rows; row++) {
                for (int column = 1; column <= columns; column++, index++) {
                    int offset = values.GetChunk(index, out byte[] kinds, out long[] bits);
                    var kind = (ExcelBufferedCellKind)kinds[offset];
                    Add(row, column, kind == ExcelBufferedCellKind.Text
                        ? ExcelBufferedCell.Restore(kind, 0L, values.GetReferenceById(bits[offset]))
                        : ExcelBufferedCell.Restore(kind, bits[offset], null));
                    if (kind == ExcelBufferedCellKind.DateTime && dateStyle != 0U) SetStyle(CellCount - 1, dateStyle);
                }
            }
        }
    }
}
