using OfficeIMO.Excel.LegacyXls.Biff;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    /// <summary>
    /// Single-pass record scanning. Each worksheet record is validated when the reader reaches
    /// it, so a row is delivered as soon as its own records are proved.
    /// </summary>
    internal sealed partial class LegacyXlsTabularDataReader {
        private void ReadWorksheetBof(ref int offset) {
            if (!(offset < _worksheetLimitOffset && TryReadDiscoveryRecord(ref offset, out RecordSlice record))) {
                throw new InvalidDataException("The XLS worksheet substream is truncated before EOF.");
            }
            ThrowIfBeyondSubstream(record);
            if (record.Type != (ushort)BiffRecordType.Bof || record.Length < 4) {
                throw new InvalidDataException("The XLS worksheet stream is missing a valid BOF record.");
            }
            ushort version = ReadUInt16(_bytes, record.PayloadOffset);
            ushort substreamType = ReadUInt16(_bytes, record.PayloadOffset + 2);
            if (version != 0x0600 || substreamType != 0x0010) {
                throw new InvalidDataException("The selected XLS substream is not a BIFF8 worksheet.");
            }
        }

        /// <summary>
        /// Moves to the next cell record, leaving it unread, or consumes the worksheet EOF.
        /// DIMENSIONS columns are recorded when they precede the first cell.
        /// </summary>
        private void LocateNextCellRow() {
            int offset = _position;
            while (true) {
                RecordSlice record = ReadWorksheetRecord(ref offset);
                if (record.Type == (ushort)BiffRecordType.Eof) {
                    _pendingRow = -1;
                    _position = offset;
                    return;
                }
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Dimensions && _previousCellRow < 0) {
                    ReadDimensions(record);
                    continue;
                }
                if (!TryGetCellBounds(record, out int row, out _, out _)) {
                    continue;
                }

                ThrowIfDecreasingRow(row);
                _pendingRow = row;
                _position = record.Offset;
                return;
            }
        }

        /// <summary>
        /// Reads the cell records of the pending row, including the String record that follows
        /// each string formula, and stops at the next row's first cell record.
        /// </summary>
        private void ReadPendingRowRecords(int currentRow) {
            int offset = _position;
            int pendingFormulaOrdinal = -1;
            ushort pendingFormulaStyle = 0;
            byte[]? bytes = _bufferedBytes;
            while (true) {
                RecordSlice record = ReadWorksheetRecord(ref offset);
                if (pendingFormulaOrdinal >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) {
                        throw MissingFormulaString();
                    }
                    StoreFormulaString(record, pendingFormulaOrdinal, pendingFormulaStyle, ref offset);
                    pendingFormulaOrdinal = -1;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Eof) {
                    _pendingRow = -1;
                    _position = offset;
                    return;
                }
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
                    continue;
                }
                if (!TryGetCellBounds(record, out int row, out _, out _)) {
                    continue;
                }
                if (row != currentRow) {
                    ThrowIfDecreasingRow(row);
                    _pendingRow = row;
                    _position = record.Offset;
                    return;
                }

                if (bytes != null) {
                    StoreBufferedCellRecord(
                        bytes,
                        currentRow,
                        record.Type,
                        record.Offset,
                        record.PayloadOffset,
                        record.Length,
                        ref pendingFormulaOrdinal,
                        ref pendingFormulaStyle);
                } else {
                    StoreCellRecord(record, ref pendingFormulaOrdinal, ref pendingFormulaStyle);
                }
            }
        }

        private Dictionary<int, string?> ReadHeaderRowRecords(int headerRow) {
            var values = new Dictionary<int, string?>();
            int offset = _position;
            int pendingHeaderColumn = -1;
            while (true) {
                int recordOffset = offset;
                RecordSlice record = ReadWorksheetRecord(ref offset);
                if (pendingHeaderColumn >= 0) {
                    if (record.Type != (ushort)BiffRecordType.String) {
                        throw MissingFormulaString();
                    }
                    values[pendingHeaderColumn] = ReadFormulaStringValue(record, ref offset);
                    pendingHeaderColumn = -1;
                    continue;
                }
                if (record.Type == (ushort)BiffRecordType.Eof) {
                    // The next locate consumes the EOF again.
                    _position = recordOffset;
                    return values;
                }
                if (record.Type == (ushort)BiffRecordType.Bof) {
                    SkipNestedChartSubstream(record, ref offset);
                    continue;
                }
                if (!TryGetCellBounds(record, out int row, out int firstColumn, out _)) {
                    continue;
                }
                if (row != headerRow) {
                    _position = record.Offset;
                    return values;
                }

                ReadHeaderCells(record, values);
                if (record.Type == (ushort)BiffRecordType.Formula && FormulaExpectsString(record)) {
                    pendingHeaderColumn = firstColumn;
                }
            }
        }

        /// <summary>
        /// Reports the column extent of the pending row's cell records without storing them;
        /// the row's values are read and validated when it is delivered.
        /// </summary>
        private void PeekPendingRowColumns(out int firstColumn, out int lastColumn) {
            firstColumn = int.MaxValue;
            lastColumn = -1;
            int offset = _position;
            while (offset < _worksheetLimitOffset && TryReadDiscoveryRecord(ref offset, out RecordSlice record)) {
                if (record.PayloadOffset + record.Length > _worksheetLimitOffset
                    || record.Type == (ushort)BiffRecordType.Eof
                    || record.Type == (ushort)BiffRecordType.Bof) {
                    return;
                }
                if (!TryGetCellBounds(record, out int row, out int recordFirstColumn, out int recordLastColumn)) {
                    continue;
                }
                if (row != _pendingRow) {
                    return;
                }
                firstColumn = Math.Min(firstColumn, recordFirstColumn);
                lastColumn = Math.Max(lastColumn, recordLastColumn);
            }
        }

        private RecordSlice ReadWorksheetRecord(ref int offset) {
            if (!(offset < _worksheetLimitOffset && TryReadDiscoveryRecord(ref offset, out RecordSlice record))) {
                throw new InvalidDataException("The XLS worksheet substream is truncated before EOF.");
            }
            ThrowIfBeyondSubstream(record);
            CheckCancellation();
            return record;
        }

        private void ReadDimensions(RecordSlice record) {
            if (record.Length < 14) return;
            int firstColumn = ReadUInt16(_bytes, record.PayloadOffset + 8);
            int columnAfterLast = ReadUInt16(_bytes, record.PayloadOffset + 10);
            // An inconsistent declaration does not widen the schema; cells still do.
            if (columnAfterLast <= firstColumn || columnAfterLast > 256) return;
            _dimensionFirstColumn = firstColumn;
            _dimensionLastColumn = columnAfterLast - 1;
        }

        private void ThrowIfDecreasingRow(int row) {
            if (row < _previousCellRow) {
                throw new InvalidDataException(
                    $"The XLS worksheet contains decreasing cell row index {row} after row {_previousCellRow}.");
            }
            _previousCellRow = row;
        }

        private void ThrowIfBeyondSubstream(RecordSlice record) {
            if (record.PayloadOffset + record.Length > _worksheetLimitOffset) {
                throw new InvalidDataException("An XLS worksheet record extends beyond its substream boundary.");
            }
        }
    }
}
