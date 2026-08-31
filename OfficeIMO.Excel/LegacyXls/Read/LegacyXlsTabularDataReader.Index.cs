using OfficeIMO.Excel.LegacyXls.Biff;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        private const int MaximumIndexedRowsPerBlock = 32;

        /// <summary>
        /// Uses BIFF8 INDEX/DBCELL metadata to locate the first cell record for each
        /// populated row without scanning the complete cell table twice. Any missing
        /// or non-canonical metadata returns false and retains the full discovery path.
        /// </summary>
        private bool TryDiscoverIndexed(
            int sheetOffset,
            int sheetEndOffset,
            int[] bufferedRowOffsets,
            out int bufferedWorksheetEndOffset,
            out int firstRow,
            out int lastRow,
            out int firstColumn,
            out int lastColumn,
            out Dictionary<int, string?>? headerValues) {
            bufferedWorksheetEndOffset = -1;
            firstRow = -1;
            lastRow = -1;
            firstColumn = 0;
            lastColumn = -1;
            headerValues = null;

            byte[] bytes = _bufferedBytes!;
            if (sheetOffset < 0
                || sheetEndOffset <= sheetOffset
                || sheetEndOffset > bytes.Length) {
                return false;
            }
            int offset = sheetOffset;
            if (!TryReadDiscoveryRecord(ref offset, out RecordSlice bof)
                || bof.PayloadOffset + bof.Length > sheetEndOffset
                || bof.Type != (ushort)BiffRecordType.Bof
                || bof.Length < 4
                || ReadDiscoveryUInt16(bof.PayloadOffset) != 0x0600
                || ReadDiscoveryUInt16(bof.PayloadOffset + 2) != 0x0010) {
                return false;
            }

            RecordSlice index = default;
            bool hasIndex = false;
            bool hasDimensions = false;
            uint dimensionFirstRow = 0;
            uint dimensionRowAfterLast = 0;
            ushort dimensionFirstColumn = 0;
            ushort dimensionColumnAfterLast = 0;
            int cellTableStartOffset = -1;

            while (offset < sheetEndOffset
                   && TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (CanCancelCurrentRead) CheckCancellation();
                if (record.PayloadOffset + record.Length > sheetEndOffset) return false;
                if (record.Type == (ushort)BiffRecordType.Index) {
                    if (hasIndex) return false;
                    index = record;
                    hasIndex = true;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Dimensions) {
                    if (hasDimensions) return false;
                    if (record.Length < 14) return false;
                    dimensionFirstRow = ReadBufferedUInt32(bytes, record.PayloadOffset);
                    dimensionRowAfterLast = ReadBufferedUInt32(bytes, record.PayloadOffset + 4);
                    dimensionFirstColumn = ReadDiscoveryUInt16(record.PayloadOffset + 8);
                    dimensionColumnAfterLast = ReadDiscoveryUInt16(record.PayloadOffset + 10);
                    hasDimensions = true;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Row || IsCellRecordType(record.Type)) {
                    cellTableStartOffset = record.Offset;
                    break;
                }
                if (record.Type == (ushort)BiffRecordType.Eof) {
                    return false;
                }
            }

            if (!hasIndex
                || index.Length < 20
                || (index.Length - 16) % sizeof(uint) != 0
                || !hasDimensions
                || cellTableStartOffset < 0
                || dimensionRowAfterLast <= dimensionFirstRow
                || dimensionRowAfterLast > bufferedRowOffsets.Length
                || dimensionColumnAfterLast <= dimensionFirstColumn
                || dimensionColumnAfterLast > 256) {
                return false;
            }

            uint indexFirstRow = ReadBufferedUInt32(bytes, index.PayloadOffset + 4);
            uint indexRowAfterLast = ReadBufferedUInt32(bytes, index.PayloadOffset + 8);
            if (indexRowAfterLast <= indexFirstRow
                || indexRowAfterLast > bufferedRowOffsets.Length
                || indexFirstRow != dimensionFirstRow
                || indexRowAfterLast != dimensionRowAfterLast) {
                return false;
            }

            int dimensionRowCount = checked((int)(dimensionRowAfterLast - dimensionFirstRow));
            Array.Clear(bufferedRowOffsets, checked((int)dimensionFirstRow), dimensionRowCount);
            int dbCellCount = (index.Length - 16) / sizeof(uint);
            int expectedDbCellCount = checked(
                (dimensionRowCount + MaximumIndexedRowsPerBlock - 1) / MaximumIndexedRowsPerBlock);
            if (dbCellCount != expectedDbCellCount) return false;
            int previousRow = -1;
            int lastDbCellEnd = -1;
            int expectedBlockStart = cellTableStartOffset;
            Span<int> blockRows = stackalloc int[MaximumIndexedRowsPerBlock];
            Span<int> blockFirstCells = stackalloc int[MaximumIndexedRowsPerBlock];

            for (int blockIndex = 0; blockIndex < dbCellCount; blockIndex++) {
                if (CanCancelCurrentRead) CheckCancellation();
                uint rawDbCellOffset = ReadBufferedUInt32(
                    bytes,
                    index.PayloadOffset + 16 + (blockIndex * sizeof(uint)));
                if (rawDbCellOffset > int.MaxValue) return false;
                int dbCellOffset = (int)rawDbCellOffset;
                if (dbCellOffset < cellTableStartOffset
                    || dbCellOffset >= sheetEndOffset
                    || dbCellOffset <= lastDbCellEnd) {
                    return false;
                }
                if (!TryReadBufferedRecordAt(bytes, dbCellOffset, out RecordSlice dbCell)
                    || dbCell.Type != (ushort)BiffRecordType.DbCell
                    || dbCell.Length < 6
                    || (dbCell.Length - sizeof(uint)) % sizeof(ushort) != 0
                    || dbCell.PayloadOffset + dbCell.Length > sheetEndOffset) {
                    return false;
                }

                int rowCount = (dbCell.Length - sizeof(uint)) / sizeof(ushort);
                int expectedRowCount = Math.Min(
                    MaximumIndexedRowsPerBlock,
                    dimensionRowCount - (blockIndex * MaximumIndexedRowsPerBlock));
                if (rowCount != expectedRowCount) return false;
                uint firstRowBackOffset = ReadBufferedUInt32(bytes, dbCell.PayloadOffset);
                if (firstRowBackOffset > dbCell.Offset) return false;
                int rowRecordOffset = checked(dbCell.Offset - (int)firstRowBackOffset);
                if (rowRecordOffset != expectedBlockStart) return false;
                int firstRowRecordEnd = -1;
                int rowRecordsEnd = rowRecordOffset;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    if (!TryReadBufferedRecordAt(bytes, rowRecordsEnd, out RecordSlice rowRecord)
                        || rowRecord.Type != (ushort)BiffRecordType.Row
                        || rowRecord.Length < 16
                        || rowRecord.PayloadOffset + rowRecord.Length > dbCell.Offset) {
                        return false;
                    }
                    int row = ReadDiscoveryUInt16(rowRecord.PayloadOffset);
                    int expectedRow = checked(
                        (int)dimensionFirstRow
                        + (blockIndex * MaximumIndexedRowsPerBlock)
                        + rowIndex);
                    if (row != expectedRow || row <= previousRow) {
                        return false;
                    }
                    blockRows[rowIndex] = row;
                    previousRow = row;
                    rowRecordsEnd = rowRecord.PayloadOffset + rowRecord.Length;
                    if (rowIndex == 0) firstRowRecordEnd = rowRecordsEnd;
                }

                int priorPosition = firstRowRecordEnd;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    ushort relativeOffset = ReadDiscoveryUInt16(
                        dbCell.PayloadOffset + sizeof(uint) + (rowIndex * sizeof(ushort)));
                    int cellOffset = checked(priorPosition + relativeOffset);
                    if (cellOffset < rowRecordsEnd
                        || cellOffset >= dbCell.Offset
                        || !TryReadBufferedRecordAt(bytes, cellOffset, out RecordSlice cell)
                        || cell.PayloadOffset + cell.Length > dbCell.Offset
                        || !TryGetCellBounds(cell, out int cellRow, out int cellFirstColumn, out int cellLastColumn)
                        || cellRow != blockRows[rowIndex]
                        || cellFirstColumn < dimensionFirstColumn
                        || cellLastColumn >= dimensionColumnAfterLast) {
                        return false;
                    }
                    blockFirstCells[rowIndex] = cellOffset;
                    priorPosition = cellOffset;
                }

                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                    bufferedRowOffsets[blockRows[rowIndex]] = blockFirstCells[rowIndex] + 1;
                }
                lastDbCellEnd = dbCell.PayloadOffset + dbCell.Length;
                expectedBlockStart = lastDbCellEnd;
            }

            if (previousRow != checked((int)dimensionRowAfterLast - 1) || lastDbCellEnd < 0) {
                return false;
            }
            offset = lastDbCellEnd;
            while (offset < sheetEndOffset
                   && TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (CanCancelCurrentRead) CheckCancellation();
                if (record.PayloadOffset + record.Length > sheetEndOffset) return false;
                if (record.Type == (ushort)BiffRecordType.Row || IsCellRecordType(record.Type)) {
                    return false;
                }
                if (record.Type != (ushort)BiffRecordType.Eof) continue;
                bufferedWorksheetEndOffset = record.Offset;
                break;
            }
            if (bufferedWorksheetEndOffset < 0) return false;

            firstRow = FindFirstIndexedRow(
                bufferedRowOffsets,
                checked((int)dimensionFirstRow),
                checked((int)dimensionRowAfterLast));
            if (firstRow < 0) return false;
            lastRow = previousRow;
            firstColumn = dimensionFirstColumn;
            lastColumn = dimensionColumnAfterLast - 1;
            if (!TryReadIndexedHeaderValues(
                    bytes,
                    bufferedRowOffsets,
                    firstRow,
                    lastRow,
                    firstColumn,
                    lastColumn,
                    bufferedWorksheetEndOffset,
                    out headerValues)) {
                return false;
            }
            return true;
        }

        private bool TryReadIndexedHeaderValues(
            byte[] bytes,
            int[] rowOffsets,
            int headerRow,
            int lastRow,
            int firstColumn,
            int lastColumn,
            int worksheetEndOffset,
            out Dictionary<int, string?> values) {
            values = new Dictionary<int, string?>();
            int offset = rowOffsets[headerRow] - 1;
            int rowEnd = worksheetEndOffset;
            for (int row = headerRow + 1; row <= lastRow; row++) {
                int marker = rowOffsets[row];
                if (marker == 0) continue;
                rowEnd = marker - 1;
                break;
            }

            int pendingFormulaColumn = -1;
            while (offset < rowEnd && TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (pendingFormulaColumn >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) return false;
                    values[pendingFormulaColumn] = ReadFormulaStringValue(record, ref offset);
                    pendingFormulaColumn = -1;
                    continue;
                }
                if (!TryGetCellBounds(
                        record,
                        out int row,
                        out int firstCellColumn,
                        out int lastCellColumn)) {
                    continue;
                }
                if (row != headerRow) return false;
                if (firstCellColumn < firstColumn || lastCellColumn > lastColumn) return false;
                ReadHeaderCells(record, values);
                if (record.Type == (ushort)BiffRecordType.Formula && FormulaExpectsString(record)) {
                    pendingFormulaColumn = firstCellColumn;
                }
            }
            return pendingFormulaColumn < 0;
        }

        private static int FindFirstIndexedRow(int[] rowOffsets, int firstRow, int rowAfterLast) {
            for (int row = firstRow; row < rowAfterLast; row++) {
                if (rowOffsets[row] != 0) return row;
            }
            return -1;
        }

        private bool TryReadBufferedRecordAt(
            byte[] bytes,
            int offset,
            out RecordSlice record) {
            int sourceLength = _bytes.Length;
            if (offset < 0 || offset > sourceLength - 4) {
                record = default;
                return false;
            }
            ushort type = (ushort)(bytes[offset] | bytes[offset + 1] << 8);
            int length = bytes[offset + 2] | bytes[offset + 3] << 8;
            int payloadOffset = offset + 4;
            if (length > sourceLength - payloadOffset) {
                record = default;
                return false;
            }
            record = new RecordSlice(type, offset, payloadOffset, length);
            return true;
        }

        private static uint ReadBufferedUInt32(byte[] bytes, int offset) =>
            (uint)(bytes[offset]
                | bytes[offset + 1] << 8
                | bytes[offset + 2] << 16
                | bytes[offset + 3] << 24);
    }
}
