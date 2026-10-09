using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    /// <summary>Forward-only physical traversal. Only the current row and explicitly requested fields are decoded.</summary>
    internal sealed class AccessNativeRowCursor : IAccessRowCursor {
        private readonly AccessNativeTable _table;
        private readonly IEnumerator<int> _pages;
        private int _page, _slot = -1, _count;
        private long _rows;
        private AccessNativeRow? _current;
        private bool _finished;
        private readonly int _filterOrdinal = -1, _filterKey;
        private readonly long _limit;
        private readonly int _valueLimit;
        private readonly bool _metadata;
        private readonly CancellationToken _cancellation;
        private bool _prefetched, _hasPrefetched;
        private bool _visible, _matched;
        internal AccessNativeRowCursor(AccessNativeTable table, CancellationToken cancellation, int filterOrdinal = -1, int filterKey = 0, long? rowLimit = null) { _table = table; _pages = table.Database.OwnedPages(table.OwnedPages, cancellation).GetEnumerator(); _filterOrdinal = filterOrdinal; _filterKey = filterKey; _limit = rowLimit ?? table.Database.MaxRows; _metadata = rowLimit.HasValue; _valueLimit = _metadata ? table.Database.MaxMetadataBytes : table.Database.MaxValueBytes; _cancellation = cancellation; }
        public bool HasRows {
            get {
                if (_filterOrdinal < 0) return _table.RowCount != 0;
                if (!_hasPrefetched && _rows == 0) { _prefetched = ReadCore(_cancellation); _hasPrefetched = true; }
                return _matched;
            }
        }
        public bool Read(CancellationToken cancellation) {
            cancellation.ThrowIfCancellationRequested();
            if (_hasPrefetched) { _hasPrefetched = false; return _visible = _prefetched; }
            return _visible = ReadCore(cancellation);
        }
        private bool ReadCore(CancellationToken cancellation) {
            if (_finished) return false;
            while (true) {
                cancellation.ThrowIfCancellationRequested();
                if (++_slot >= _count) {
                    bool found = false;
                    while (_pages.MoveNext()) {
                        cancellation.ThrowIfCancellationRequested(); _page = _pages.Current; OfficeByteView page = _table.Database.Page(_page);
                        if (page[0] != 1) continue;
                        if (I32(page, 4) != _table.DefinitionPage) throw new InvalidDataException("Native Access table usage map refers to another table's data page.");
                        _count = U16(page, _table.Database.Layout.DataRowCount); if (_count > 255) throw new InvalidDataException("Native Access data-page row count is invalid.");
                        _slot = 0; found = true; break;
                    }
                    if (!found) { _finished = true; _current = null; if (_rows != _table.RowCount) throw new InvalidDataException("Native Access table row count disagrees with its owned live records."); return false; }
                    if (_count == 0) continue;
                }
                OfficeByteView data = _table.Database.Page(_page, 1); int flags = U16(data, _table.Database.Layout.DataRowDirectory + _slot * 2);
                if ((flags & 0x8000) != 0) continue;
                if (_rows == _limit) throw new InvalidDataException("Native Access reader exceeds its row limit.");
                _rows++; _current = new AccessNativeRow(_table, _table.Database.Row(_page, _slot, true, cancellation), _valueLimit, _metadata);
                if (_filterOrdinal >= 0 && !Equals(_current.Value(_filterOrdinal, cancellation), _filterKey)) continue;
                _matched = true; return true;
            }
        }
        internal AccessNativeRow Current => _visible && _current != null ? _current : throw new InvalidOperationException("Call Read before accessing a current row.");
        public bool IsSpecified(int ordinal) { Current.ValidateOrdinal(ordinal); return true; }
        public bool IsNull(int ordinal) => Current.IsNull(ordinal);
        public Stream OpenBinary(int ordinal, CancellationToken cancellation) => Current.OpenBinary(ordinal, cancellation);
        public object? GetValue(int ordinal, CancellationToken cancellation) => Current.Value(ordinal, cancellation);
        public void Dispose() { _finished = true; _current = null; _pages.Dispose(); }
    }

    internal sealed class AccessNativeRow {
        private readonly AccessNativeTable _table;
        private readonly OfficeByteView _data;
        private readonly int _columnCount, _nullBytes, _variableCount, _valuesEnd;
        private readonly int _fieldStart;
        private readonly int[]? _jet3Offsets;
        private readonly object?[] _values;
        private readonly bool[] _decoded;
        private readonly int _valueLimit;
        private readonly bool _metadata;
        internal AccessNativeRow(AccessNativeTable table, OfficeByteView data, int valueLimit, bool metadata) {
            _table = table; _data = data; _valueLimit = valueLimit; _metadata = metadata;
            _fieldStart = table.Database.Layout.RowColumnCountSize;
            _columnCount = _fieldStart == 1 ? Slice(data, 0, 1)[0] : U16(data, 0);
            if (_columnCount > table.MaxColumns) throw new InvalidDataException("Native Access row declares fields outside its table schema.");
            _nullBytes = (_columnCount + 7) / 8;
            if (data.Length < _fieldStart + _nullBytes) throw new InvalidDataException("Native Access row is truncated before its null mask.");
            int countPosition = data.Length - _nullBytes - _fieldStart;
            _variableCount = table.MaxVariableColumns == 0 ? 0 : _fieldStart == 1 ? Slice(data, countPosition, 1)[0] : U16(data, countPosition);
            if (_variableCount > table.MaxVariableColumns) throw new InvalidDataException("Native Access row variable-field count exceeds its table schema.");
            if (_fieldStart == 1 && table.MaxVariableColumns != 0) {
                _jet3Offsets = Jet3Offsets(data, countPosition, _variableCount);
                _valuesEnd = _jet3Offsets[_variableCount];
            } else _valuesEnd = data.Length - _nullBytes - (table.MaxVariableColumns == 0 ? 0 : 4 + _variableCount * 2);
            if (_valuesEnd < _fieldStart) throw new InvalidDataException("Native Access row offset directory overlaps its field data.");
            _values = new object?[table.Columns.Count]; _decoded = new bool[table.Columns.Count];
        }
        private static int[] Jet3Offsets(OfficeByteView data, int countPosition, int variableCount) {
            // Byte offsets are extended by ordered column indices at each 256-byte boundary.
            // Physical jump slots depend on row length; trailing 0xff slots are padding.
            int jumpSlots = (data.Length - 1) / 256;
            int lastOffset = countPosition - jumpSlots - 1, directoryStart = lastOffset - variableCount;
            if (directoryStart < 1) throw new InvalidDataException("Jet3 variable directory overlaps its row header.");
            var jumps = new List<int>(); bool padding = false;
            for (int slot = 0; slot < jumpSlots; slot++) {
                int index = data[countPosition - slot - 1];
                if (index == 255 && (variableCount < 255 || (slot + 1) * 256 > directoryStart)) { padding = true; continue; }
                if (padding || index > variableCount || jumps.Count != 0 && index < jumps[jumps.Count - 1])
                    throw new InvalidDataException("Jet3 jump indices are invalid or out of order.");
                jumps.Add(index);
            }
            var offsets = new int[variableCount + 1]; int segment = 0;
            for (int index = 0; index < offsets.Length; index++) {
                while (segment < jumps.Count && jumps[segment] <= index) segment++;
                int value = data[lastOffset - index] + segment * 256;
                if (value < 1 || value > directoryStart || index != 0 && value < offsets[index - 1])
                    throw new InvalidDataException("Jet3 variable offsets exceed their row data or are out of order.");
                offsets[index] = value;
            }
            return offsets;
        }
        internal void ValidateOrdinal(int ordinal) { if ((uint)ordinal >= (uint)_table.Columns.Count) throw new IndexOutOfRangeException(); }
        internal byte[] NativeBytes() { _table.Database.AccountMetadata(_data.Length); return _data.ToArray(); }
        internal bool IsNull(int ordinal) {
            ValidateOrdinal(ordinal); AccessNativeColumn column = _table.Columns[ordinal];
            return column.Type != 1 && !Present(column);
        }
        private bool Present(AccessNativeColumn column) => column.Number < _columnCount && (_data[_data.Length - _nullBytes + column.Number / 8] & (1 << (column.Number % 8))) != 0;
        private OfficeByteView Payload(AccessNativeColumn column) {
            int start, length;
            if (column.Variable) {
                if (column.VariableIndex >= _variableCount) throw new InvalidDataException("Native Access non-null field lacks a variable offset.");
                if (_jet3Offsets != null) { start = _jet3Offsets[column.VariableIndex]; length = _jet3Offsets[column.VariableIndex + 1] - start; }
                else { int position = _data.Length - _nullBytes - 4 - column.VariableIndex * 2; start = U16(_data, position); length = U16(_data, position - 2) - start; }
            } else { start = checked(_fieldStart + column.FixedOffset); length = column.Size; }
            if (start < _fieldStart || length < 0 || start > _valuesEnd - length) throw new InvalidDataException("Native Access field overlaps the row directory or lies outside its record.");
            return Slice(_data, start, length);
        }
        internal Stream OpenBinary(int ordinal, CancellationToken cancellation) {
            ValidateOrdinal(ordinal); _table.Database.Document.EnsureNotDisposed(); cancellation.ThrowIfCancellationRequested();
            AccessNativeColumn column = _table.Columns[ordinal];
            if (IsNull(ordinal)) throw new InvalidCastException("The current native field is null.");
            if (column.Type != 9 && column.Type != 11 && column.Type != 17 || column.Calculated) throw new InvalidCastException("The current native field is not binary.");
            OfficeByteView payload = Payload(column);
            if (column.Type == 11) return new AccessNativeLongValueStream(_table.Database, payload, cancellation);
            if (payload.Length > _table.Database.MaxValueBytes) throw new InvalidDataException("Native Access binary value exceeds MaxValueBytes.");
            return new OfficeIMO.Core.Internal.OfficeDocumentReadStream(new MemoryStream(payload.ToArray(), writable: false), _table.Database.Document.EnsureNotDisposed, cancellation);
        }
        internal object? Value(int ordinal, CancellationToken cancellation) {
            ValidateOrdinal(ordinal); cancellation.ThrowIfCancellationRequested(); if (_decoded[ordinal]) return _values[ordinal];
            AccessNativeColumn column = _table.Columns[ordinal]; bool present = Present(column);
            object? value = null;
            if (column.Type == 1 && !column.Calculated) value = present;
            else if (present) {
                value = _table.Database.DecodeScalar(column, Payload(column), cancellation, _metadata ? _valueLimit : (int?)null);
            }
            _decoded[ordinal] = true; return _values[ordinal] = value;
        }
    }
}
