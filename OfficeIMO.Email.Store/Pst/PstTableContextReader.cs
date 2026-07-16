using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstTableContextReader {
    private readonly PstHeap _heap;
    private readonly bool _isUnicode;
    private readonly EmailStoreReaderOptions _options;
    private readonly CancellationToken _cancellationToken;

    internal PstTableContextReader(PstHeap heap, bool isUnicode, EmailStoreReaderOptions options,
        CancellationToken cancellationToken) {
        _heap = heap;
        _isUnicode = isUnicode;
        _options = options;
        _cancellationToken = cancellationToken;
    }

    internal IReadOnlyList<IReadOnlyList<MapiProperty>> ReadRows() {
        if (_heap.ClientSignature != 0x7C) throw new InvalidDataException("The PST node is not a Table Context.");
        byte[] info = _heap.GetAllocation(_heap.UserRoot);
        if (info.Length < 22 || info[0] != 0x7C) throw new InvalidDataException("The PST TCINFO header is invalid.");

        int columnCount = info[1];
        int rowSize = PstBinary.UInt16(info, 8);
        int existenceOffset = PstBinary.UInt16(info, 6);
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
            if (dataOffset < 0 || dataSize < 0 || dataOffset > rowSize - dataSize) {
                throw new InvalidDataException("A PST table column points outside its row.");
            }
            columns.Add(new PstTableColumn(tag, dataOffset, dataSize, bitIndex));
        }

        IReadOnlyList<int> rowIndexes = ReadRowIndexes(rowIndexHid);
        PstDataTree rowData = _heap.ResolveHnidTree(rowsHnid, _options.MaxDecodedPropertyBytesPerItem);
        var rows = new List<IReadOnlyList<MapiProperty>>(rowIndexes.Count);
        foreach (int rowIndex in rowIndexes) {
            _cancellationToken.ThrowIfCancellationRequested();
            byte[] row = GetRow(rowData.Blocks, rowSize, rowIndex);
            var properties = new List<MapiProperty>(columns.Count);
            long decodedBytes = 0;
            foreach (PstTableColumn column in columns) {
                int existenceByte = existenceOffset + column.BitIndex / 8;
                if (existenceByte >= row.Length ||
                    (row[existenceByte] & (1 << (7 - column.BitIndex % 8))) == 0) continue;
                MapiProperty property = DecodeColumn(row, column, ref decodedBytes);
                properties.Add(property);
            }
            rows.Add(properties);
        }
        return rows;
    }

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

    private MapiProperty DecodeColumn(byte[] row, PstTableColumn column, ref long decodedBytes) {
        ushort id = (ushort)(column.Tag >> 16);
        var type = (MapiPropertyType)(ushort)column.Tag;
        object? value;
        byte[]? rawData = null;
        int offset = column.DataOffset;
        switch (type) {
            case MapiPropertyType.Integer16:
                value = PstBinary.Int16(row, offset);
                break;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                value = PstBinary.Int32(row, offset);
                break;
            case MapiPropertyType.Floating32:
                value = BitConverter.ToSingle(row, offset);
                break;
            case MapiPropertyType.Boolean:
                value = row[offset] != 0;
                break;
            case MapiPropertyType.Integer64:
            case MapiPropertyType.Currency:
                value = PstBinary.Int64(row, offset);
                break;
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
                value = BitConverter.ToDouble(row, offset);
                break;
            case MapiPropertyType.Time:
                try {
                    value = new DateTimeOffset(DateTime.FromFileTimeUtc(PstBinary.Int64(row, offset)));
                } catch (ArgumentOutOfRangeException) {
                    value = null;
                }
                break;
            default:
                uint hnid = PstBinary.UInt32(row, offset);
                rawData = _heap.ResolveHnid(hnid, _options.MaxDecodedPropertyBytesPerItem);
                decodedBytes = checked(decodedBytes + rawData.Length);
                if (decodedBytes > _options.MaxDecodedPropertyBytesPerItem) {
                    throw new EmailStoreLimitExceededException(
                        nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem), decodedBytes,
                        _options.MaxDecodedPropertyBytesPerItem);
                }
                value = PstPropertyContextReader.DecodeVariable(type, rawData);
                break;
        }
        return new MapiProperty(id, type, value) { RawData = rawData };
    }

    private static byte[] GetRow(IReadOnlyList<byte[]> blocks, int rowSize, int requestedIndex) {
        int index = requestedIndex;
        foreach (byte[] block in blocks) {
            int rowsInBlock = block.Length / rowSize;
            if (index < rowsInBlock) {
                int offset = checked(index * rowSize);
                var row = new byte[rowSize];
                Buffer.BlockCopy(block, offset, row, 0, rowSize);
                return row;
            }
            index -= rowsInBlock;
        }
        throw new InvalidDataException("A PST table row index points outside the Row Matrix.");
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
}
