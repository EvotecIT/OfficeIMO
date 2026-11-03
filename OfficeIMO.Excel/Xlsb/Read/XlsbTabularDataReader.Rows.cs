using OfficeIMO.Excel.Xlsb.Biff12;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Single-pass record scanning. Each record is framed, budgeted and validated once, when
    /// the reader reaches it, so a row is delivered as soon as its own records are proved.
    /// </summary>
    internal sealed partial class XlsbTabularDataReader {
        private readonly int[] _rowSpanBounds = new int[32];
        private int _rowSpanCount;
        private int _currentRowIndex = -1;
        private int _firstRecordType = -1;
        private int _lastRecordType = -1;
        private int _dimensionFirstColumn;
        private int _dimensionLastColumn = -1;
        private bool _sawDimension;
        private bool _sawBeginSheetData;
        private bool _inSheetData;
        private bool _sawEndSheetData;
        private bool _pendingRowLocated;

        /// <summary>
        /// Moves to the next row header followed by at least one cell record, validating every
        /// record before it. Empty rows between populated rows are later delivered as blank rows;
        /// trailing empty rows are not. At the end of the sheet data, the remaining records of the
        /// part are validated before the reader reports that no row remains.
        /// </summary>
        private void LocatePendingRow() {
            _pendingRowLocated = true;
            _hasPendingRow = false;
            byte[] bytes = _records.Buffer;
            int length = _records.Length;
            int position = _records.Position;
            while (true) {
                if (position == length) {
                    _records.Position = position;
                    FinishPart();
                    return;
                }

                int recordOffset = position;
                int payloadOffset = ReadRecordFrame(bytes, position, length, out int recordType, out int recordSize);
                position = payloadOffset + recordSize;
                CountRecord(recordType);
                if (_inSheetData) {
                    if (recordType == BrtRowHdr) {
                        int rowIndex = ValidateRowHeader(
                            new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize),
                            _rowSpanBounds,
                            out _rowSpanCount);
                        if (_currentRowIndex >= 0 && rowIndex <= _currentRowIndex) {
                            throw new InvalidDataException(
                                $"The XLSB worksheet contains non-increasing row index {rowIndex} after header row {_currentRowIndex}.");
                        }

                        _currentRowIndex = rowIndex;
                        if (RowHasCells(bytes, position, length)) {
                            _pendingRowIndex = rowIndex;
                            _hasPendingRow = true;
                            _records.Position = position;
                            return;
                        }

                        continue;
                    }

                    if (IsCellRecord(recordType)) {
                        // An empty row header was proved to precede no cell record, so a cell
                        // here can only precede the first row header.
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {recordOffset} appears before a row header.");
                    }
                }

                ProcessStructureRecord(recordType, recordOffset, payloadOffset, recordSize, bytes);
            }
        }

        /// <summary>
        /// Reads the cells of the located row into the current-row arrays. The loop stops at the
        /// next row header or sheet-data end without consuming it.
        /// </summary>
        private void ReadPendingRowCells() {
            byte[] bytes = _records.Buffer;
            int length = _records.Length;
            int position = _records.Position;
            int previousCellColumn = -1;
            int fieldCount = _headers.Length;
            int firstColumn = _firstColumn;
            bool converted = _options.CellValueConverter != null;
            while (position < length) {
                int recordOffset = position;
                int payloadOffset = ReadRecordFrame(bytes, position, length, out int recordType, out int recordSize);
                if (recordType == BrtRowHdr || recordType == BrtEndSheetData) {
                    break;
                }

                position = payloadOffset + recordSize;
                CountRecord(recordType);
                if (!IsCellRecord(recordType)) {
                    ProcessStructureRecord(recordType, recordOffset, payloadOffset, recordSize, bytes);
                    continue;
                }

                _cellBudget.Consume();
                int column = ValidateCellRecord(bytes, recordType, recordOffset, payloadOffset, recordSize);
                ValidateCellPlacement(column, recordOffset, ref previousCellColumn);
                int ordinal = column - firstColumn;
                if ((uint)ordinal >= (uint)fieldCount) {
                    ThrowCellOutsideSchema(column);
                }

                if (converted) {
                    StoreCell(new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize), ordinal);
                    continue;
                }

                switch (recordType) {
                    case BrtCellRk:
                        StoreValidatedRkCell(bytes, payloadOffset, ordinal);
                        break;
                    case BrtCellReal:
                        StoreValidatedRealCell(bytes, payloadOffset, ordinal);
                        break;
                    case BrtCellIsst:
                        StoreValidatedSharedStringCell(bytes, payloadOffset, ordinal);
                        break;
                    default:
                        StoreCellFast(bytes, recordType, payloadOffset, ordinal);
                        break;
                }
            }

            _records.Position = position;
            if (position == length) {
                // The part ended inside sheet data; report the same structural failure.
                FinishPart();
            }
        }

        /// <summary>
        /// Reads the header row with the same record and cell validation as data rows.
        /// </summary>
        private Dictionary<int, string?> ReadHeaderRowCells() {
            var values = new Dictionary<int, string?>();
            byte[] bytes = _records.Buffer;
            int length = _records.Length;
            int position = _records.Position;
            int previousCellColumn = -1;
            while (position < length) {
                int recordOffset = position;
                int payloadOffset = ReadRecordFrame(bytes, position, length, out int recordType, out int recordSize);
                if (recordType == BrtRowHdr || recordType == BrtEndSheetData) {
                    break;
                }

                position = payloadOffset + recordSize;
                CountRecord(recordType);
                if (!IsCellRecord(recordType)) {
                    ProcessStructureRecord(recordType, recordOffset, payloadOffset, recordSize, bytes);
                    continue;
                }

                _cellBudget.Consume();
                int column = ValidateCellRecord(bytes, recordType, recordOffset, payloadOffset, recordSize);
                ValidateCellPlacement(column, recordOffset, ref previousCellColumn);
                DecodedCell cell = DecodeCell(new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize));
                values[cell.Column] = CellToHeaderText(cell);
            }

            _records.Position = position;
            if (position == length) {
                FinishPart();
            }

            return values;
        }

        /// <summary>
        /// Reports the column extent of the located row's framed cell records without
        /// consuming or validating their payloads; the row is validated when it is read.
        /// </summary>
        private void PeekPendingRowColumns(out int firstColumn, out int lastColumn) {
            firstColumn = int.MaxValue;
            lastColumn = -1;
            byte[] bytes = _records.Buffer;
            int length = _records.Length;
            int position = _records.Position;
            while (position < length) {
                int payloadOffset = ReadRecordFrame(bytes, position, length, out int recordType, out int recordSize);
                if (recordType == BrtRowHdr || recordType == BrtEndSheetData) {
                    return;
                }

                position = payloadOffset + recordSize;
                if (!IsCellRecord(recordType) || recordSize < sizeof(int)) {
                    continue;
                }

                int column = Unsafe.ReadUnaligned<int>(ref bytes[payloadOffset]);
                if (column < 0 || column >= ExcelA1.MaxColumns) {
                    continue;
                }

                firstColumn = Math.Min(firstColumn, column);
                lastColumn = Math.Max(lastColumn, column);
            }
        }

        private bool RowHasCells(byte[] bytes, int position, int length) {
            while (position < length) {
                int payloadOffset = ReadRecordFrame(bytes, position, length, out int recordType, out int recordSize);
                if (IsCellRecord(recordType)) {
                    return true;
                }
                if (recordType == BrtRowHdr || recordType == BrtEndSheetData) {
                    return false;
                }

                position = payloadOffset + recordSize;
            }

            return false;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void CountRecord(int recordType) {
            _recordBudget.Consume();
            if (_firstRecordType < 0) {
                _firstRecordType = recordType;
            }
            _lastRecordType = recordType;
            _recordsSinceCancellationCheck++;
            if ((_recordsSinceCancellationCheck & 1023) == 0) {
                ThrowIfReadCancellationRequested();
            }
        }

        /// <summary>
        /// Applies the worksheet structure rules to records that are neither rows nor cells
        /// inside sheet data.
        /// </summary>
        private void ProcessStructureRecord(
            int recordType,
            int recordOffset,
            int payloadOffset,
            int recordSize,
            byte[] bytes) {
            switch (recordType) {
                case BrtWsDim:
                    if (_sawDimension) {
                        throw new InvalidDataException(
                            "The XLSB worksheet contains more than one BrtWsDim record.");
                    }
                    if (_sawBeginSheetData) {
                        throw new InvalidDataException(
                            "The XLSB worksheet contains a misplaced BrtWsDim record.");
                    }

                    ReadWorksheetDimension(
                        new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize),
                        out _dimensionFirstColumn,
                        out _dimensionLastColumn);
                    _sawDimension = true;
                    return;
                case BrtBeginSheetData:
                    if (_sawBeginSheetData) {
                        throw new InvalidDataException(
                            "The XLSB worksheet contains duplicate or nested BrtBeginSheetData records.");
                    }
                    if (!_sawDimension) {
                        throw new InvalidDataException(
                            "The XLSB worksheet is missing the required BrtWsDim record before BrtBeginSheetData.");
                    }

                    _sawBeginSheetData = true;
                    _inSheetData = true;
                    return;
                case BrtEndSheetData:
                    if (!_inSheetData) {
                        throw new InvalidDataException(
                            "The XLSB worksheet contains BrtEndSheetData without a matching BrtBeginSheetData record.");
                    }

                    _sawEndSheetData = true;
                    _inSheetData = false;
                    _currentRowIndex = -1;
                    return;
                default:
                    if (!_inSheetData && (recordType == BrtRowHdr || IsCellRecord(recordType))) {
                        throw new InvalidDataException(
                            $"The XLSB row or cell record at offset {recordOffset} appears outside BrtBeginSheetData/BrtEndSheetData.");
                    }

                    return;
            }
        }

        private void FinishPart() {
            ThrowIfReadCancellationRequested();
            if (_firstRecordType != BrtBeginSheet
                || _lastRecordType != BrtEndSheet) {
                throw new InvalidDataException(
                    "The XLSB worksheet is missing its outer BrtBeginSheet/BrtEndSheet boundaries.");
            }
            if (!_sawDimension) {
                throw new InvalidDataException(
                    "The XLSB worksheet is missing the required BrtWsDim record.");
            }
            if (!_sawBeginSheetData) {
                throw new InvalidDataException(
                    "The XLSB worksheet does not contain the required BrtBeginSheetData record.");
            }
            if (!_sawEndSheetData) {
                throw new InvalidDataException(
                    "The XLSB worksheet ended before the required BrtEndSheetData record.");
            }
        }

        /// <summary>
        /// Frames one BIFF12 record header at <paramref name="position"/>, which must be inside
        /// the part, and returns its payload offset.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int ReadRecordFrame(
            byte[] bytes,
            int position,
            int length,
            out int recordType,
            out int recordSize) {
            int recordOffset = position;
            ref byte data = ref MemoryMarshal.GetReference(bytes.AsSpan());
            int firstTypeByte = Unsafe.Add(ref data, position++);
            recordType = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                if (position == length) {
                    ThrowTruncatedHeader("record type");
                }
                int secondTypeByte = Unsafe.Add(ref data, position++);
                recordType |= (secondTypeByte & 0x7F) << 7;
                if (recordType < 128) {
                    ThrowNonCanonicalRecordType();
                }
            }

            if (position == length) {
                ThrowTruncatedHeader("record size");
            }
            int sizeByte = Unsafe.Add(ref data, position++);
            recordSize = sizeByte & 0x7F;
            if ((sizeByte & 0x80) != 0) {
                recordSize = ReadLongRecordSize(bytes, ref position, length, recordSize);
            }
            if (recordSize > _limits.MaxRecordBytes) {
                ThrowRecordTooLarge(recordOffset, recordSize);
            }
            if (recordSize > length - position) {
                ThrowRecordBeyondPart(recordOffset, recordSize, length - position);
            }

            return position;
        }

        private static int ReadLongRecordSize(byte[] bytes, ref int position, int length, int recordSize) {
            for (int sizeIndex = 1; sizeIndex < 4; sizeIndex++) {
                if (position == length) {
                    ThrowTruncatedHeader("record size");
                }
                int sizeByte = bytes[position++];
                recordSize |= (sizeByte & 0x7F) << (sizeIndex * 7);
                if ((sizeByte & 0x80) == 0) {
                    return recordSize;
                }
            }

            throw new InvalidDataException(
                "The BIFF12 record size header is invalid.");
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowTruncatedHeader(string fieldName) =>
            throw new EndOfStreamException(
                $"The BIFF12 stream ended inside the {fieldName} header.");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowNonCanonicalRecordType() =>
            throw new InvalidDataException(
                "The BIFF12 record type uses a non-canonical two-byte encoding.");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowRecordTooLarge(int recordOffset, int recordSize) =>
            throw new InvalidDataException(
                $"The BIFF12 record at offset {recordOffset} declares {recordSize} payload bytes, exceeding the configured limit of {_limits.MaxRecordBytes} bytes.");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowRecordBeyondPart(int recordOffset, int recordSize, int remaining) =>
            throw new EndOfStreamException(
                $"The BIFF12 record at offset {recordOffset} declares {recordSize} payload bytes but only {remaining} remain.");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void ThrowCellOutsideSchema(int column) =>
            throw new InvalidDataException(
                $"The XLSB row {_pendingRowIndex + 1} contains column {column + 1} outside the {FieldCount} column(s) " +
                "established by its header row, first data row or worksheet dimension. " +
                "Set ExcelReaderOptions.A1Range to read a wider range.");
    }
}
