using OfficeIMO.Excel.Xlsb.Biff12;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private static void DiscoverDataColumns(
            Stream worksheetPart,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            XlsbCellReadBudget cellBudget,
            int styleCount,
            int sharedStringCount,
            CancellationToken cancellationToken,
            out int firstColumn,
            out int lastColumn,
            out int firstDataRow,
            out int lastDataRow) {
            if (!worksheetPart.CanSeek) {
                throw new InvalidOperationException(
                    "XLSB reads require a seekable worksheet part for schema discovery.");
            }

            firstColumn = int.MaxValue;
            lastColumn = -1;
            firstDataRow = -1;
            lastDataRow = -1;
            int currentRow = -1;
            int previousCellColumn = -1;
            long startPosition = worksheetPart.Position;
            try {
                using var scanner = new XlsbStreamRecordSliceReader(
                    worksheetPart,
                    limits.MaxRecordBytes,
                    recordBudget,
                    leaveOpen: true);
                bool inSheetData = false;
                bool sawBeginSheetData = false;
                bool sawEndSheetData = false;
                bool sawDimension = false;
                int firstRecordType = -1;
                int lastRecordType = -1;
                int recordsSinceCancellationCheck = 0;
                var currentRowSpanBounds = new int[32];
                int currentRowSpanCount = 0;
                cancellationToken.ThrowIfCancellationRequested();
                while (scanner.TryRead(out XlsbRecordSlice record)) {
                    if (firstRecordType < 0) {
                        firstRecordType = record.Type;
                    }
                    lastRecordType = record.Type;
                    recordsSinceCancellationCheck++;
                    if ((recordsSinceCancellationCheck & 1023) == 0) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    if (record.Type == BrtWsDim) {
                        if (sawDimension) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains more than one BrtWsDim record.");
                        }
                        if (sawBeginSheetData) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains a misplaced BrtWsDim record.");
                        }

                        ReadWorksheetDimension(record, out _, out _);
                        sawDimension = true;
                        continue;
                    }
                    if (record.Type == BrtBeginSheetData) {
                        if (sawBeginSheetData) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains duplicate or nested BrtBeginSheetData records.");
                        }
                        if (!sawDimension) {
                            throw new InvalidDataException(
                                "The XLSB worksheet is missing the required BrtWsDim record before BrtBeginSheetData.");
                        }

                        sawBeginSheetData = true;
                        inSheetData = true;
                        continue;
                    }
                    if (record.Type == BrtEndSheetData) {
                        if (!inSheetData) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains BrtEndSheetData without a matching BrtBeginSheetData record.");
                        }

                        sawEndSheetData = true;
                        inSheetData = false;
                        currentRow = -1;
                        continue;
                    }
                    if (!inSheetData) {
                        if (record.Type == BrtRowHdr || IsCellRecord(record.Type)) {
                            throw new InvalidDataException(
                                $"The XLSB row or cell record at offset {record.RecordOffset} appears outside BrtBeginSheetData/BrtEndSheetData.");
                        }

                        continue;
                    }
                    if (record.Type == BrtRowHdr) {
                        int nextRow = ValidateRowHeader(
                            record,
                            currentRowSpanBounds,
                            out currentRowSpanCount);
                        if (currentRow >= 0 && nextRow <= currentRow) {
                            throw new InvalidDataException(
                                $"The XLSB worksheet contains non-increasing row index {nextRow} after header row {currentRow}.");
                        }

                        currentRow = nextRow;
                        previousCellColumn = -1;
                        continue;
                    }
                    if (!IsCellRecord(record.Type)) {
                        continue;
                    }
                    if (currentRow < 0) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {record.RecordOffset} appears before a row header.");
                    }

                    cellBudget.Consume();
                    if (firstDataRow < 0) {
                        firstDataRow = currentRow;
                    }
                    lastDataRow = currentRow;
                    int column = ValidateCellPayloadStructure(
                        record,
                        styleCount,
                        sharedStringCount,
                        limits.MaxStringCharacters);
                    if (column < 0 || column >= A1.MaxColumns) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
                    }
                    if (column <= previousCellColumn) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {record.RecordOffset} is duplicated or out of order within its row.");
                    }
                    previousCellColumn = column;
                    bool covered = false;
                    for (int index = 0; index < currentRowSpanCount; index++) {
                        int offset = index * 2;
                        if (currentRowSpanBounds[offset] <= column
                            && column <= currentRowSpanBounds[offset + 1]) {
                            covered = true;
                            break;
                        }
                    }
                    if (!covered) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {record.RecordOffset} for column {column} is not covered by its BrtRowHdr column spans.");
                    }

                    firstColumn = Math.Min(firstColumn, column);
                    lastColumn = Math.Max(lastColumn, column);
                }

                cancellationToken.ThrowIfCancellationRequested();
                if (firstRecordType != BrtBeginSheet
                    || lastRecordType != BrtEndSheet) {
                    throw new InvalidDataException(
                        "The XLSB worksheet is missing its outer BrtBeginSheet/BrtEndSheet boundaries.");
                }
                if (!sawDimension) {
                    throw new InvalidDataException(
                        "The XLSB worksheet is missing the required BrtWsDim record.");
                }
                if (!sawBeginSheetData) {
                    throw new InvalidDataException(
                        "The XLSB worksheet does not contain the required BrtBeginSheetData record.");
                }
                if (!sawEndSheetData) {
                    throw new InvalidDataException(
                        "The XLSB worksheet ended before the required BrtEndSheetData record.");
                }
            } finally {
                worksheetPart.Position = startPosition;
            }
        }

        private static void ReadWorksheetDimension(
            XlsbRecordSlice record,
            out int firstColumn,
            out int lastColumn) {
            if (record.Size != 16) {
                throw new InvalidDataException(
                    $"The BrtWsDim record at offset {record.RecordOffset} has invalid payload length {record.Size}.");
            }

            var cursor = record.CreateCursor();
            uint firstRow = cursor.ReadUInt32();
            uint lastRow = cursor.ReadUInt32();
            uint firstColumnValue = cursor.ReadUInt32();
            uint lastColumnValue = cursor.ReadUInt32();
            if (firstRow > lastRow
                || lastRow >= A1.MaxRows
                || firstColumnValue > lastColumnValue
                || lastColumnValue >= A1.MaxColumns) {
                throw new InvalidDataException(
                    $"The BrtWsDim record at offset {record.RecordOffset} contains an invalid worksheet range.");
            }

            firstColumn = checked((int)firstColumnValue);
            lastColumn = checked((int)lastColumnValue);
        }

        private static int ValidateCellPayloadStructure(
            XlsbRecordSlice record,
            int styleCount,
            int sharedStringCount,
            int maxStringCharacters) {
            try {
                var cursor = record.CreateCursor();
                int column = cursor.ReadInt32();
                uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
                if (styleIndex >= styleCount) {
                    throw new InvalidDataException(
                        $"The XLSB cell record at offset {record.RecordOffset} refers to missing cell format " +
                        $"{styleIndex}; the styles part exposes {styleCount} format(s).");
                }
                switch (record.Type) {
                    case BrtCellBlank:
                        break;
                    case BrtCellRk:
                        cursor.ReadUInt32();
                        break;
                    case BrtCellIsst:
                        uint sharedStringIndex = cursor.ReadUInt32();
                        if (sharedStringIndex >= sharedStringCount) {
                            throw new InvalidDataException(
                                $"The XLSB cell record at offset {record.RecordOffset} refers to missing shared string " +
                                $"{sharedStringIndex}; the shared-string part exposes {sharedStringCount} item(s).");
                        }
                        break;
                    case BrtCellError:
                    case BrtCellBool:
                        cursor.ReadByte();
                        break;
                    case BrtCellReal:
                        cursor.ReadDouble();
                        break;
                    case BrtCellSt:
                        cursor.ReadWideString(maxStringCharacters);
                        break;
                    case BrtCellRString:
                        cursor.ReadByte();
                        cursor.ReadWideString(maxStringCharacters);
                        break;
                    case BrtFmlaString:
                        cursor.ReadWideString(maxStringCharacters);
                        ValidateFormulaPayloadTail(record, ref cursor);
                        break;
                    case BrtFmlaNum:
                        cursor.ReadDouble();
                        ValidateFormulaPayloadTail(record, ref cursor);
                        break;
                    case BrtFmlaBool:
                    case BrtFmlaError:
                        cursor.ReadByte();
                        ValidateFormulaPayloadTail(record, ref cursor);
                        break;
                    default:
                        throw new InvalidOperationException(
                            $"Unsupported XLSB cell record type {record.Type}.");
                }

                return column;
            } catch (EndOfStreamException exception) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} is truncated.",
                    exception);
            }
        }
    }
}
