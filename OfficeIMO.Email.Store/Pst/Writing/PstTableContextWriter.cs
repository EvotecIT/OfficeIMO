using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstWriterTableRow {
    internal PstWriterTableRow(uint rowId, IEnumerable<MapiProperty> properties) {
        RowId = rowId;
        Properties = properties == null
            ? Array.Empty<MapiProperty>()
            : properties.Where(item => item != null).ToArray();
    }

    internal uint RowId { get; }
    internal IReadOnlyList<MapiProperty> Properties { get; }
}

internal static class PstTableContextWriter {
    private const int MaximumBlockPayload = 8176;

    internal static PstWriterContextResult Write(PstWriterFile file,
        IEnumerable<PstWriterTableRow> sourceRows, int codePage,
        IEnumerable<MapiProperty>? requiredColumns,
        Action<EmailStoreDiagnostic>? reportDiagnostic, string location) {
        if (file == null) throw new ArgumentNullException(nameof(file));
        if (sourceRows == null) sourceRows = Array.Empty<PstWriterTableRow>();
        if (requiredColumns == null) {
            throw new ArgumentNullException(nameof(requiredColumns),
                "Streaming PST tables require an explicit column contract.");
        }
        Column[] columns = CreateColumns(requiredColumns);
        if (columns.Length > byte.MaxValue) {
            throw new NotSupportedException("A PST Table Context cannot contain more than 255 columns.");
        }

        AssignLayout(columns, out int dataEnd, out int twoByteEnd, out int oneByteEnd);
        int bitmapBytes = (columns.Length + 7) / 8;
        int rowSize = checked(oneByteEnd + bitmapBytes);
        if (rowSize <= 0 || rowSize > MaximumBlockPayload) {
            throw new NotSupportedException("A PST table row exceeds the supported Unicode block size.");
        }
        int rowsPerFullBlock = MaximumBlockPayload / rowSize;
        if (rowsPerFullBlock <= 0) throw new NotSupportedException("A PST row is larger than one data block.");

        var heap = new PstWriterHeap(0x7C);
        var info = new byte[22 + columns.Length * 8];
        uint infoHid = heap.Add(info);
        var rowIndexHeader = new byte[8];
        uint rowIndexHeaderHid = heap.Add(rowIndexHeader);
        const uint rowMatrixNid = 0x3F;
        uint nextLocalIndex = 2;
        string matrixPath = file.CreateTemporaryIndexPath("table-matrix");
        string rowIndexPath = file.CreateTemporaryIndexPath("table-row-index");
        string subnodePath = file.CreateTemporaryIndexPath("table-subnodes");
        int rowCount = 0;
        ulong rowMatrixBid = 0;
        ulong subnodeBid;
        using (var subnodes = new PstWriterSubnodeJournal(subnodePath))
        try {
            using (var matrix = new FileStream(matrixPath, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
            using (var rowIndex = new FileStream(rowIndexPath, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.Read, 64 * 1024, FileOptions.SequentialScan)) {
                var matrixBlock = new byte[MaximumBlockPayload];
                int rowsInBlock = 0;
                uint previousRowId = 0;
                bool hasPreviousRow = false;
                foreach (PstWriterTableRow row in sourceRows) {
                    if (hasPreviousRow && row.RowId <= previousRowId) {
                        throw new InvalidDataException(
                            "Streaming PST table rows must have unique ascending row identifiers.");
                    }
                    previousRowId = row.RowId;
                    hasPreviousRow = true;
                    var values = row.Properties.GroupBy(item => item.PropertyId)
                        .ToDictionary(group => group.Key, group => group.Last());
                    values[0x67F2] = new MapiProperty(0x67F2, MapiPropertyType.Integer32,
                        unchecked((int)row.RowId));
                    if (!values.ContainsKey(0x67F3)) {
                        values[0x67F3] = new MapiProperty(0x67F3, MapiPropertyType.Integer32, 0);
                    }

                    int rowOffset = rowsInBlock * rowSize;
                    foreach (Column column in columns) {
                        if (!values.TryGetValue(column.PropertyId, out MapiProperty? property)) continue;
                        try {
                            EncodedCell cell = EncodeCell(property, column, codePage);
                            int destination = rowOffset + column.DataOffset;
                            if (cell.Bytes != null) {
                                uint hnid = 0;
                                if (cell.Bytes.Length > 0) {
                                    uint nid = checked((nextLocalIndex++ << 5) | 0x1FU);
                                    subnodes.Add(new PstWriterSubnode(nid, file.WriteDataTree(cell.Bytes)));
                                    hnid = nid;
                                }
                                PstBinary.WriteUInt32(matrixBlock, destination, hnid);
                            } else {
                                Buffer.BlockCopy(cell.InlineBytes!, 0, matrixBlock,
                                    destination, cell.InlineBytes!.Length);
                            }
                            int bitmapOffset = rowOffset + oneByteEnd + column.BitIndex / 8;
                            matrixBlock[bitmapOffset] |= checked((byte)(1 << (7 - column.BitIndex % 8)));
                        } catch (Exception exception) when (exception is ArgumentException ||
                            exception is InvalidCastException || exception is FormatException ||
                            exception is OverflowException || exception is NotSupportedException) {
                            reportDiagnostic?.Invoke(new EmailStoreDiagnostic(
                                "EMAIL_STORE_PST_WRITE_TABLE_CELL_OMITTED",
                                string.Concat("Table property 0x",
                                    property.PropertyTag.ToString("X8", CultureInfo.InvariantCulture),
                                    " was omitted: ", exception.Message),
                                EmailStoreDiagnosticSeverity.Warning, location));
                        }
                    }

                    var indexRecord = new byte[8];
                    PstBinary.WriteUInt32(indexRecord, 0, row.RowId);
                    PstBinary.WriteUInt32(indexRecord, 4, checked((uint)rowCount));
                    rowIndex.Write(indexRecord, 0, indexRecord.Length);
                    rowCount = checked(rowCount + 1);
                    rowsInBlock++;
                    if (rowsInBlock == rowsPerFullBlock) {
                        matrix.Write(matrixBlock, 0, matrixBlock.Length);
                        Array.Clear(matrixBlock, 0, matrixBlock.Length);
                        rowsInBlock = 0;
                    }
                }
                if (rowsInBlock > 0) matrix.Write(matrixBlock, 0, rowsInBlock * rowSize);

                if (rowCount > 0) {
                    matrix.Position = 0;
                    rowMatrixBid = file.WriteDataTree(matrix, matrix.Length);
                }
                rowIndex.Position = 0;
                PstWriterBth.Complete(heap, rowIndexHeader, 4, 4, rowIndex, rowCount);
            }

            info[0] = 0x7C;
            info[1] = checked((byte)columns.Length);
            PstBinary.WriteUInt16(info, 2, dataEnd);
            PstBinary.WriteUInt16(info, 4, twoByteEnd);
            PstBinary.WriteUInt16(info, 6, oneByteEnd);
            PstBinary.WriteUInt16(info, 8, rowSize);
            PstBinary.WriteUInt32(info, 10, rowIndexHeaderHid);
            PstBinary.WriteUInt32(info, 14, rowCount == 0 ? 0 : rowMatrixNid);
            PstBinary.WriteUInt32(info, 18, 0);
            for (int index = 0; index < columns.Length; index++) {
                Column column = columns[index];
                int offset = 22 + index * 8;
                PstBinary.WriteUInt32(info, offset,
                    ((uint)column.PropertyId << 16) | (ushort)column.PropertyType);
                PstBinary.WriteUInt16(info, offset + 4, column.DataOffset);
                info[offset + 6] = checked((byte)column.Size);
                info[offset + 7] = checked((byte)column.BitIndex);
            }

            ulong dataBid = file.WriteDataTreeBlocks(heap.Build(infoHid));
            int subnodeCount = subnodes.Count + (rowCount == 0 ? 0 : 1);
            subnodeBid = PstWriterSubnodeTree.WriteSorted(file,
                EnumerateTableSubnodes(rowCount, rowMatrixNid, rowMatrixBid, subnodes), subnodeCount);
            return new PstWriterContextResult(dataBid, subnodeBid);
        } finally {
            TryDelete(matrixPath);
            TryDelete(rowIndexPath);
        }
    }

    private static IEnumerable<PstWriterSubnode> EnumerateTableSubnodes(int rowCount,
        uint rowMatrixNid, ulong rowMatrixBid, PstWriterSubnodeJournal values) {
        if (rowCount > 0) yield return new PstWriterSubnode(rowMatrixNid, rowMatrixBid);
        foreach (PstWriterSubnode value in values.ReadAll()) yield return value;
    }

    private static void TryDelete(string path) {
        try { if (File.Exists(path)) File.Delete(path); }
        catch (IOException) { }
        catch (UnauthorizedAccessException) { }
    }

    private static Column[] CreateColumns(IEnumerable<MapiProperty> requiredColumns) {
        var types = new Dictionary<ushort, MapiPropertyType> {
            [0x67F2] = MapiPropertyType.Integer32,
            [0x67F3] = MapiPropertyType.Integer32
        };
        foreach (MapiProperty property in requiredColumns) types[property.PropertyId] = property.PropertyType;
        Column[] columns = types.OrderBy(item => ((uint)item.Key << 16) | (ushort)item.Value)
            .Select(item => new Column(item.Key, item.Value, GetTableWidth(item.Value))).ToArray();
        int nextBit = 2;
        foreach (Column column in columns) {
            column.BitIndex = column.PropertyId == 0x67F2 ? 0 : column.PropertyId == 0x67F3 ? 1 : nextBit++;
        }
        return columns;
    }

    private static void AssignLayout(Column[] columns, out int dataEnd,
        out int twoByteEnd, out int oneByteEnd) {
        Column rowId = columns.Single(item => item.PropertyId == 0x67F2);
        Column rowVersion = columns.Single(item => item.PropertyId == 0x67F3);
        rowId.DataOffset = 0;
        rowVersion.DataOffset = 4;
        int cursor = 8;
        foreach (Column column in columns.Where(item => item != rowId && item != rowVersion && item.Size >= 4)
            .OrderByDescending(item => item.Size).ThenBy(item => item.PropertyId)) {
            int alignment = column.Size >= 8 ? 8 : 4;
            cursor = (cursor + alignment - 1) & ~(alignment - 1);
            column.DataOffset = cursor;
            cursor += column.Size;
        }
        dataEnd = cursor;
        foreach (Column column in columns.Where(item => item.Size == 2).OrderBy(item => item.PropertyId)) {
            cursor = (cursor + 1) & ~1;
            column.DataOffset = cursor;
            cursor += 2;
        }
        twoByteEnd = cursor;
        foreach (Column column in columns.Where(item => item.Size == 1).OrderBy(item => item.PropertyId)) {
            column.DataOffset = cursor++;
        }
        oneByteEnd = cursor;
    }

    private static EncodedCell EncodeCell(MapiProperty property, Column column, int codePage) {
        switch (column.PropertyType) {
            case MapiPropertyType.Integer16:
                return EncodedCell.Inline(BitConverter.GetBytes(
                    Convert.ToInt16(property.Value ?? 0, CultureInfo.InvariantCulture)));
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
                return EncodedCell.Inline(BitConverter.GetBytes(
                    Convert.ToInt32(property.Value ?? 0, CultureInfo.InvariantCulture)));
            case MapiPropertyType.Floating32:
                return EncodedCell.Inline(BitConverter.GetBytes(
                    Convert.ToSingle(property.Value ?? 0F, CultureInfo.InvariantCulture)));
            case MapiPropertyType.Boolean:
                return EncodedCell.Inline(new[] {
                    Convert.ToBoolean(property.Value ?? false, CultureInfo.InvariantCulture) ? (byte)1 : (byte)0 });
            case MapiPropertyType.Integer64:
            case MapiPropertyType.Currency:
                return EncodedCell.Inline(BitConverter.GetBytes(
                    Convert.ToInt64(property.Value ?? 0L, CultureInfo.InvariantCulture)));
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
                return EncodedCell.Inline(BitConverter.GetBytes(
                    Convert.ToDouble(property.Value ?? 0D, CultureInfo.InvariantCulture)));
            case MapiPropertyType.Time:
                return EncodedCell.Inline(PstPropertyValueWriter.EncodeVariable(property, codePage));
            default:
                return EncodedCell.Variable(PstPropertyValueWriter.EncodeVariable(property, codePage));
        }
    }

    private static int GetTableWidth(MapiPropertyType type) {
        switch (type) {
            case MapiPropertyType.Boolean: return 1;
            case MapiPropertyType.Integer16: return 2;
            case MapiPropertyType.Integer32:
            case MapiPropertyType.ErrorCode:
            case MapiPropertyType.Floating32: return 4;
            case MapiPropertyType.Integer64:
            case MapiPropertyType.Currency:
            case MapiPropertyType.Floating64:
            case MapiPropertyType.FloatingTime:
            case MapiPropertyType.Time: return 8;
            default: return 4;
        }
    }

    private sealed class Column {
        internal Column(ushort propertyId, MapiPropertyType propertyType, int size) {
            PropertyId = propertyId;
            PropertyType = propertyType;
            Size = size;
        }
        internal ushort PropertyId { get; }
        internal MapiPropertyType PropertyType { get; }
        internal int Size { get; }
        internal int DataOffset { get; set; }
        internal int BitIndex { get; set; }
    }

    private sealed class EncodedCell {
        private EncodedCell(byte[]? inlineBytes, byte[]? bytes) {
            InlineBytes = inlineBytes;
            Bytes = bytes;
        }
        internal byte[]? InlineBytes { get; }
        internal byte[]? Bytes { get; }
        internal static EncodedCell Inline(byte[] bytes) => new EncodedCell(bytes, null);
        internal static EncodedCell Variable(byte[] bytes) => new EncodedCell(null, bytes);
    }
}
