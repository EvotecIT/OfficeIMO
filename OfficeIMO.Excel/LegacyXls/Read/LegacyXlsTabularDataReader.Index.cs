using OfficeIMO.Excel.LegacyXls.Biff;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        private const int MaximumIndexedRowsPerBlock = 32;
        // Blocks validated before the first Read: the header row and the row after it must be indexed.
        private const int EagerlyValidatedIndexedBlocks = 2;

        private int _indexPayloadOffset;
        private int _indexedCellTableStart;
        private int _indexedSheetEnd;
        private int _indexedFirstRow;
        private int _indexedRowCount;
        private int _indexedBlockCount;
        private int _indexedFirstColumn;
        private int _indexedColumnAfterLast;
        private int _indexedValidatedBlocks;
        private int _indexedNextBlockStart;
        // First row whose block is not validated yet; int.MaxValue once every block is validated or the
        // reader does not use INDEX metadata.
        private int _indexedValidatedRowEnd = int.MaxValue;
        private int[]? _indexedBlockRows;
        private int[]? _indexedBlockFirstCells;

        /// <summary>
        /// Uses BIFF8 INDEX/DBCELL metadata to locate the first cell record for each
        /// populated row without scanning the complete cell table twice. Any missing
        /// or non-canonical metadata returns false and retains the full discovery path.
        /// Opening validates the DBCELL chain, the records after it and the first blocks;
        /// each later block's ROW records and cell pointers are validated just before its
        /// rows are read, while they are about to be decoded anyway.
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
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
                    continue;
                }
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

            _indexPayloadOffset = index.PayloadOffset;
            _indexedCellTableStart = cellTableStartOffset;
            _indexedSheetEnd = sheetEndOffset;
            _indexedFirstRow = checked((int)dimensionFirstRow);
            _indexedRowCount = dimensionRowCount;
            _indexedBlockCount = dbCellCount;
            _indexedFirstColumn = dimensionFirstColumn;
            _indexedColumnAfterLast = dimensionColumnAfterLast;

            // The DBCELL records alone prove that the blocks tile the cell table in row order.
            int expectedBlockStart = cellTableStartOffset;
            for (int blockIndex = 0; blockIndex < dbCellCount; blockIndex++) {
                if (CanCancelCurrentRead) CheckCancellation();
                if (!TryReadIndexedDbCell(bytes, blockIndex, expectedBlockStart, out RecordSlice dbCell, out _)) {
                    return false;
                }
                expectedBlockStart = dbCell.PayloadOffset + dbCell.Length;
            }

            offset = expectedBlockStart;
            while (offset < sheetEndOffset
                   && TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (CanCancelCurrentRead) CheckCancellation();
                if (record.PayloadOffset + record.Length > sheetEndOffset) return false;
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Row || IsCellRecordType(record.Type)) {
                    return false;
                }
                if (record.Type != (ushort)BiffRecordType.Eof) continue;
                bufferedWorksheetEndOffset = record.Offset;
                break;
            }
            if (bufferedWorksheetEndOffset < 0) return false;

            _indexedBlockRows = new int[MaximumIndexedRowsPerBlock];
            _indexedBlockFirstCells = new int[MaximumIndexedRowsPerBlock];
            _indexedValidatedBlocks = 0;
            _indexedNextBlockStart = cellTableStartOffset;
            _indexedValidatedRowEnd = _indexedFirstRow;
            int eagerBlocks = Math.Min(EagerlyValidatedIndexedBlocks, dbCellCount);
            for (int blockIndex = 0; blockIndex < eagerBlocks; blockIndex++) {
                if (!TryValidateNextIndexedBlock(bytes, bufferedRowOffsets)) return false;
            }

            // Every indexed row has a ROW record and a first cell, so the dimensions bound the rows.
            firstRow = _indexedFirstRow;
            lastRow = checked((int)dimensionRowAfterLast - 1);
            firstColumn = dimensionFirstColumn;
            lastColumn = dimensionColumnAfterLast - 1;
            return TryReadIndexedHeaderValues(
                bytes,
                bufferedRowOffsets,
                firstRow,
                Math.Min(lastRow, _indexedValidatedRowEnd - 1),
                firstColumn,
                lastColumn,
                bufferedWorksheetEndOffset,
                out headerValues);
        }

        private bool TryReadIndexedDbCell(
            byte[] bytes,
            int blockIndex,
            int expectedBlockStart,
            out RecordSlice dbCell,
            out int rowRecordOffset) {
            rowRecordOffset = -1;
            uint rawDbCellOffset = ReadBufferedUInt32(
                bytes,
                _indexPayloadOffset + 16 + (blockIndex * sizeof(uint)));
            // INDEX holds workbook stream offsets; a worksheet slice starts at _streamOffsetBase.
            long relativeDbCellOffset = (long)rawDbCellOffset - _streamOffsetBase;
            if (relativeDbCellOffset < _indexedCellTableStart
                || relativeDbCellOffset >= _indexedSheetEnd
                || relativeDbCellOffset < expectedBlockStart) {
                dbCell = default;
                return false;
            }
            int dbCellOffset = (int)relativeDbCellOffset;
            if (!TryReadBufferedRecordAt(bytes, dbCellOffset, out dbCell)
                || dbCell.Type != (ushort)BiffRecordType.DbCell
                || dbCell.Length < 6
                || (dbCell.Length - sizeof(uint)) % sizeof(ushort) != 0
                || dbCell.PayloadOffset + dbCell.Length > _indexedSheetEnd) {
                return false;
            }

            int rowCount = (dbCell.Length - sizeof(uint)) / sizeof(ushort);
            int expectedRowCount = Math.Min(
                MaximumIndexedRowsPerBlock,
                _indexedRowCount - (blockIndex * MaximumIndexedRowsPerBlock));
            if (rowCount != expectedRowCount) return false;
            uint firstRowBackOffset = ReadBufferedUInt32(bytes, dbCell.PayloadOffset);
            if (firstRowBackOffset > dbCell.Offset) return false;
            rowRecordOffset = checked(dbCell.Offset - (int)firstRowBackOffset);
            return rowRecordOffset == expectedBlockStart;
        }

        /// <summary>Validates the next block's ROW records and cell pointers and records its row offsets.</summary>
        private bool TryValidateNextIndexedBlock(byte[] bytes, int[] bufferedRowOffsets) {
            int blockIndex = _indexedValidatedBlocks;
            if (!TryReadIndexedDbCell(bytes, blockIndex, _indexedNextBlockStart, out RecordSlice dbCell, out int rowRecordOffset)) {
                return false;
            }
            int[] blockRows = _indexedBlockRows!;
            int[] blockFirstCells = _indexedBlockFirstCells!;
            int rowCount = (dbCell.Length - sizeof(uint)) / sizeof(ushort);
            int firstRowRecordEnd = -1;
            int rowRecordsEnd = rowRecordOffset;
            int expectedRow = checked(_indexedFirstRow + (blockIndex * MaximumIndexedRowsPerBlock));
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++, expectedRow++) {
                if (!TryReadBufferedRecordAt(bytes, rowRecordsEnd, out RecordSlice rowRecord)
                    || rowRecord.Type != (ushort)BiffRecordType.Row
                    || rowRecord.Length < 16
                    || rowRecord.PayloadOffset + rowRecord.Length > dbCell.Offset
                    || ReadDiscoveryUInt16(rowRecord.PayloadOffset) != expectedRow) {
                    return false;
                }
                blockRows[rowIndex] = expectedRow;
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
                    || cellFirstColumn < _indexedFirstColumn
                    || cellLastColumn >= _indexedColumnAfterLast) {
                    return false;
                }
                blockFirstCells[rowIndex] = cellOffset;
                priorPosition = cellOffset;
            }

            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                bufferedRowOffsets[blockRows[rowIndex]] = blockFirstCells[rowIndex] + 1;
            }
            _indexedNextBlockStart = dbCell.PayloadOffset + dbCell.Length;
            _indexedValidatedBlocks = blockIndex + 1;
            _indexedValidatedRowEnd = _indexedValidatedBlocks == _indexedBlockCount
                ? int.MaxValue
                : _indexedFirstRow + (_indexedValidatedBlocks * MaximumIndexedRowsPerBlock);
            return true;
        }

        /// <summary>
        /// Validates indexed blocks through <paramref name="row"/>. A block that fails validation
        /// switches the reader to full discovery, which must describe the same table.
        /// </summary>
        private void EnsureIndexedRowsValidated(int row) {
            byte[] bytes = _bufferedBytes!;
            int[] rowOffsets = _bufferedRowOffsets!;
            while (row >= _indexedValidatedRowEnd) {
                if (CanCancelCurrentRead) CheckCancellation();
                bool validated;
                try {
                    validated = TryValidateNextIndexedBlock(bytes, rowOffsets);
                } catch (InvalidDataException) {
                    validated = false;
                }
                if (!validated) {
                    FallBackToFullDiscovery(rowOffsets);
                    return;
                }
            }
        }

        private void FallBackToFullDiscovery(int[] rowOffsets) {
            _indexedValidatedRowEnd = int.MaxValue;
            _usedIndexedDiscovery = false;
            Discover(
                _worksheetStartOffset,
                rowOffsets,
                out int worksheetEndOffset,
                out int firstRow,
                out int lastRow,
                out int firstColumn,
                out int lastColumn,
                out Dictionary<int, string?>? headerValues);
            int firstDataRow = _hasHeaderRow && firstRow >= 0 ? checked(firstRow + 1) : firstRow;
            int fieldCount = lastColumn >= firstColumn ? checked(lastColumn - firstColumn + 1) : 0;
            if (firstDataRow != _firstDataRow
                || lastRow != _lastDataRow
                || firstColumn != _firstColumn
                || fieldCount != FieldCount
                || !BuildHeaders(fieldCount, firstColumn, _hasHeaderRow, headerValues, _options).AsSpan().SequenceEqual(_headers)) {
                throw new InvalidDataException(
                    "The XLS worksheet INDEX and DBCELL records describe a different table than its cell records.");
            }
            _bufferedWorksheetEndOffset = worksheetEndOffset;
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
            while (offset < rowEnd) {
                if (rowEnd - offset < sizeof(uint)) return false;
                if (!TryReadDiscoveryRecord(ref offset, out RecordSlice record)) break;
                if (offset > rowEnd) return false;
                if (pendingFormulaColumn >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) return false;
                    values[pendingFormulaColumn] = ReadFormulaStringValue(record, ref offset);
                    pendingFormulaColumn = -1;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
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
            return offset <= rowEnd && pendingFormulaColumn < 0;
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
