using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstTableContextReader {
    private readonly PstHeap _heap;
    private readonly bool _isUnicode;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly Action<string>? _reportCellWarning;
    private readonly long _maximumDecodedPropertyBytes;

    internal PstTableContextReader(PstHeap heap, bool isUnicode, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, Action<string>? reportCellWarning = null,
        long? maximumDecodedPropertyBytes = null) {
        _heap = heap;
        _isUnicode = isUnicode;
        _options = options;
        _cancellationToken = cancellationToken;
        _reportCellWarning = reportCellWarning;
        _maximumDecodedPropertyBytes = maximumDecodedPropertyBytes ?? options.MaxDecodedPropertyBytesPerItem;
    }

    internal IReadOnlyList<IReadOnlyList<MapiProperty>> ReadRows() => EnumerateRows().ToArray();

    /// <summary>Streams table rows while retaining only the active row-matrix data block.</summary>
    internal IEnumerable<IReadOnlyList<MapiProperty>> EnumerateRows() {
        if (_heap.ClientSignature != 0x7C) throw new InvalidDataException("The PST node is not a Table Context.");
        byte[] info = _heap.GetAllocation(_heap.UserRoot);
        if (info.Length < 22 || info[0] != 0x7C) throw new InvalidDataException("The PST TCINFO header is invalid.");

        int columnCount = info[1];
        int rowSize = PstBinary.UInt16(info, 8);
        int existenceBytes = checked((columnCount + 7) / 8);
        int existenceOffset = rowSize - existenceBytes;
        uint rowIndexHid = PstBinary.UInt32(info, 10);
        uint rowsHnid = PstBinary.UInt32(info, 14);
        if (rowSize <= 0 || existenceOffset < 0 || existenceOffset > rowSize ||
            22 + checked(columnCount * 8) > info.Length) {
            throw new InvalidDataException("The PST TCINFO row layout is invalid.");
        }

        var columns = new List<PstTableColumn>(columnCount);
        for (int index = 0; index < columnCount; index++) {
            int offset = 22 + index * 8;
            uint tag = PstBinary.UInt32(info, offset);
            int dataOffset = PstBinary.UInt16(info, offset + 4);
            int dataSize = info[offset + 6];
            int bitIndex = info[offset + 7];
            if (bitIndex >= columnCount || dataOffset < 0 || dataSize < 0 ||
                dataOffset > existenceOffset - dataSize) {
                throw new InvalidDataException("A PST table column points outside its row.");
            }
            columns.Add(new PstTableColumn(tag, dataOffset, dataSize, bitIndex));
        }

        IReadOnlyList<int> rowIndexes;
        try {
            rowIndexes = ReadRowIndexes(rowIndexHid);
        } catch (InvalidDataException exception) {
            throw CreateTableInfoException(info, columnCount, rowSize, existenceOffset,
                rowIndexHid, rowsHnid, "row-index", exception);
        }
        IEnumerable<byte[]> rowBlocks = _heap.EnumerateHnidBlocks(
            rowsHnid, Math.Min(_options.MaxDecodedTableBytes, _maximumDecodedPropertyBytes));
        long decodedBytes = 0;
        using (var cursor = new PstTableRowCursor(rowBlocks, rowSize)) {
            foreach (int rowIndex in rowIndexes.OrderBy(index => index)) {
                _cancellationToken.ThrowIfCancellationRequested();
                PstTableRow row;
                try {
                    row = cursor.GetRow(rowIndex);
                } catch (InvalidDataException exception) {
                    throw CreateTableInfoException(info, columnCount, rowSize, existenceOffset,
                        rowIndexHid, rowsHnid, "row-matrix", exception);
                }
                var properties = new List<MapiProperty>(columns.Count);
                foreach (PstTableColumn column in columns) {
                    int existenceByte = row.Offset + existenceOffset + column.BitIndex / 8;
                    if (existenceByte >= row.Offset + row.Length ||
                        (row.Bytes[existenceByte] & (1 << (7 - column.BitIndex % 8))) == 0) continue;
                    MapiProperty? property = DecodeColumn(row, rowIndex, column, ref decodedBytes);
                    if (property != null) properties.Add(property);
                }
                yield return properties;
            }
        }
    }

    private static InvalidDataException CreateTableInfoException(
        byte[] info, int columnCount, int rowSize, int existenceOffset,
        uint rowIndexHid, uint rowsHnid, string component, InvalidDataException exception) =>
        new InvalidDataException(string.Concat(
            "TCINFO ", component,
            " bytes=", info.Length.ToString(CultureInfo.InvariantCulture),
            " columns=", columnCount.ToString(CultureInfo.InvariantCulture),
            " row-size=", rowSize.ToString(CultureInfo.InvariantCulture),
            " existence-offset=", existenceOffset.ToString(CultureInfo.InvariantCulture),
            " row-index=0x", rowIndexHid.ToString("X8", CultureInfo.InvariantCulture),
            " rows=0x", rowsHnid.ToString("X8", CultureInfo.InvariantCulture),
            ": ", exception.Message), exception);

    private IReadOnlyList<int> ReadRowIndexes(uint headerHid) {
        if (headerHid == 0) return Array.Empty<int>();
        byte[] header = _heap.GetAllocation(headerHid);
        int valueSize = _isUnicode ? 4 : 2;
        if (header.Length < 8 || header[0] != 0xB5 || header[1] != 4 || header[2] != valueSize) {
            throw new InvalidDataException("The PST Table Context row-index BTH is invalid.");
        }
        int levels = header[3];
        uint root = PstBinary.UInt32(header, 4);
        var indexes = new List<int>();
        foreach (byte[] record in _heap.EnumerateBthLeafRecords(root, 4, valueSize, levels)) {
            int index = _isUnicode ? PstBinary.Int32(record, 4) : PstBinary.UInt16(record, 4);
            if (index < 0) throw new InvalidDataException("A PST table row index is negative.");
            indexes.Add(index);
        }
        return indexes;
    }

    private MapiProperty? DecodeColumn(
        PstTableRow row, int rowIndex, PstTableColumn column, ref long decodedBytes) {
        ushort id = (ushort)(column.Tag >> 16);
        var type = (MapiPropertyType)(ushort)column.Tag;
        object? value;
        byte[]? rawData = null;
        int offset = row.Offset + column.DataOffset;
        switch (type) {
            case MapiPropertyType.Integer16:
                value = PstBinary.Int16(row.Bytes, offset);
                break;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                value = PstBinary.Int32(row.Bytes, offset);
                break;
            case MapiPropertyType.Floating32:
                value = BitConverter.ToSingle(row.Bytes, offset);
                break;
            case MapiPropertyType.Boolean:
                value = row.Bytes[offset] != 0;
                break;
            case MapiPropertyType.Currency:
                value = PstBinary.Int64(row.Bytes, offset) / 10000m;
                break;
            case MapiPropertyType.Integer64:
                value = PstBinary.Int64(row.Bytes, offset);
                break;
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
                value = BitConverter.ToDouble(row.Bytes, offset);
                break;
            case MapiPropertyType.Time:
                try {
                    value = new DateTimeOffset(DateTime.FromFileTimeUtc(PstBinary.Int64(row.Bytes, offset)));
                } catch (ArgumentOutOfRangeException) {
                    value = null;
                }
                break;
            default:
                uint hnid = PstBinary.UInt32(row.Bytes, offset);
                try {
                    rawData = _heap.ResolveHnid(hnid, _maximumDecodedPropertyBytes);
                } catch (InvalidDataException exception) {
                    _reportCellWarning?.Invoke(string.Concat(
                        "Table row ", rowIndex.ToString(CultureInfo.InvariantCulture),
                        " column tag 0x", column.Tag.ToString("X8", CultureInfo.InvariantCulture),
                        " type 0x", ((ushort)type).ToString("X4", CultureInfo.InvariantCulture),
                        " width ", column.DataSize.ToString(CultureInfo.InvariantCulture),
                        " contains invalid HNID 0x", hnid.ToString("X8", CultureInfo.InvariantCulture),
                        ": ", exception.Message));
                    return null;
                }
                decodedBytes = checked(decodedBytes + rawData.Length);
                if (decodedBytes > _maximumDecodedPropertyBytes) {
                    throw new EmailStoreLimitExceededException(
                        nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), decodedBytes,
                        _maximumDecodedPropertyBytes);
                }
                value = PstPropertyContextReader.DecodeVariable(type, rawData);
                break;
        }
        return new MapiProperty(id, type, value) { RawData = rawData };
    }

    private sealed class PstTableColumn {
        internal PstTableColumn(uint tag, int dataOffset, int dataSize, int bitIndex) {
            Tag = tag;
            DataOffset = dataOffset;
            DataSize = dataSize;
            BitIndex = bitIndex;
        }

        internal uint Tag { get; }
        internal int DataOffset { get; }
        internal int DataSize { get; }
        internal int BitIndex { get; }
    }

    private sealed class PstTableRowCursor : IDisposable {
        private readonly IEnumerator<byte[]> _blocks;
        private readonly int _rowSize;
        private byte[]? _currentBlock;
        private int _currentBlockStart;
        private int _currentBlockRows;
        private int _lastRequestedIndex = -1;

        internal PstTableRowCursor(IEnumerable<byte[]> blocks, int rowSize) {
            _blocks = blocks.GetEnumerator();
            _rowSize = rowSize;
        }

        internal PstTableRow GetRow(int requestedIndex) {
            if (requestedIndex < _lastRequestedIndex) {
                throw new InvalidDataException("PST table row indexes must be read in ascending matrix order.");
            }
            _lastRequestedIndex = requestedIndex;
            while (_currentBlock == null ||
                   requestedIndex >= _currentBlockStart + _currentBlockRows) {
                _currentBlockStart += _currentBlockRows;
                if (!_blocks.MoveNext()) {
                    throw new InvalidDataException("A PST table row index points outside the Row Matrix.");
                }
                _currentBlock = _blocks.Current;
                _currentBlockRows = _currentBlock.Length / _rowSize;
                if (_currentBlockRows == 0) {
                    throw new InvalidDataException("A PST Row Matrix block is smaller than one table row.");
                }
            }

            int localIndex = requestedIndex - _currentBlockStart;
            int offset = checked(localIndex * _rowSize);
            return new PstTableRow(_currentBlock, offset, _rowSize);
        }

        public void Dispose() => _blocks.Dispose();
    }

    private readonly struct PstTableRow {
        internal PstTableRow(byte[] bytes, int offset, int length) {
            Bytes = bytes;
            Offset = offset;
            Length = length;
        }

        internal byte[] Bytes { get; }
        internal int Offset { get; }
        internal int Length { get; }
    }
}
