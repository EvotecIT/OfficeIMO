using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Package;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private const int DiscoveryBudgetBatchSize = 1024;

        private static void DiscoverDataColumns(
            XlsbPooledPartStream worksheetPart,
            XlsbImportOptions limits,
            XlsbRecordReadBudget recordBudget,
            XlsbCellReadBudget cellBudget,
            int styleCount,
            int sharedStringCount,
            bool useCachedFormulaResult,
            CancellationToken cancellationToken,
            out int firstColumn,
            out int lastColumn,
            out int firstDataRow,
            out int lastDataRow,
            out XlsbValidatedRowPlan? validatedRowPlan) {
            firstColumn = int.MaxValue;
            lastColumn = -1;
            firstDataRow = -1;
            lastDataRow = -1;
            validatedRowPlan = null;
            int currentRow = -1;
            int previousCellColumn = -1;
            byte[] bytes = worksheetPart.Buffer;
            int dataLength = worksheetPart.DataLength;
            int position = 0;
            {
                bool inSheetData = false;
                bool sawBeginSheetData = false;
                bool sawEndSheetData = false;
                bool sawDimension = false;
                int firstRecordType = -1;
                int lastRecordType = -1;
                int recordsSinceCancellationCheck = 0;
                int pendingRecordBudget = 0;
                int pendingCellBudget = 0;
                var currentRowSpanBounds = new int[32];
                var rowPlanBuilder = new XlsbValidatedRowPlanBuilder();
                int currentRowSpanCount = 0;
                cancellationToken.ThrowIfCancellationRequested();
                // Discovery is the full-buffer validation pass. Fuse BIFF12 framing with the
                // common fixed-cell checks so large worksheets are validated once without a
                // method dispatch and record-object path for every cell.
                while (position < dataLength) {
                    int recordOffset = position;
                    int firstTypeByte = bytes[position++];
                    int recordType = firstTypeByte & 0x7F;
                    if ((firstTypeByte & 0x80) != 0) {
                        if (position == dataLength) {
                            throw new EndOfStreamException(
                                "The BIFF12 stream ended inside the record type header.");
                        }
                        int secondTypeByte = bytes[position++];
                        recordType |= (secondTypeByte & 0x7F) << 7;
                        if (recordType < 128) {
                            throw new InvalidDataException(
                                "The BIFF12 record type uses a non-canonical two-byte encoding.");
                        }
                    }

                    if (position == dataLength) {
                        throw new EndOfStreamException(
                            "The BIFF12 stream ended inside the record size header.");
                    }
                    int sizeByte = bytes[position++];
                    int recordSize = sizeByte & 0x7F;
                    if ((sizeByte & 0x80) != 0) {
                        bool complete = false;
                        for (int sizeIndex = 1; sizeIndex < 4; sizeIndex++) {
                            if (position == dataLength) {
                                throw new EndOfStreamException(
                                    "The BIFF12 stream ended inside the record size header.");
                            }
                            sizeByte = bytes[position++];
                            recordSize |= (sizeByte & 0x7F) << (sizeIndex * 7);
                            if ((sizeByte & 0x80) == 0) {
                                complete = true;
                                break;
                            }
                        }
                        if (!complete) {
                            throw new InvalidDataException(
                                "The BIFF12 record size header is invalid.");
                        }
                    }
                    if (recordSize > limits.MaxRecordBytes) {
                        throw new InvalidDataException(
                            $"The BIFF12 record at offset {recordOffset} declares {recordSize} payload bytes, exceeding the configured limit of {limits.MaxRecordBytes} bytes.");
                    }
                    if (recordSize > dataLength - position) {
                        throw new EndOfStreamException(
                            $"The BIFF12 record at offset {recordOffset} declares {recordSize} payload bytes but only {dataLength - position} remain.");
                    }
                    int payloadOffset = position;
                    position += recordSize;

                    pendingRecordBudget++;
                    if (pendingRecordBudget == DiscoveryBudgetBatchSize) {
                        recordBudget.Consume(pendingRecordBudget);
                        pendingRecordBudget = 0;
                    }
                    if (firstRecordType < 0) {
                        firstRecordType = recordType;
                    }
                    lastRecordType = recordType;
                    recordsSinceCancellationCheck++;
                    if ((recordsSinceCancellationCheck & 1023) == 0) {
                        cancellationToken.ThrowIfCancellationRequested();
                    }

                    if (recordType == BrtWsDim) {
                        if (sawDimension) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains more than one BrtWsDim record.");
                        }
                        if (sawBeginSheetData) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains a misplaced BrtWsDim record.");
                        }

                        ReadWorksheetDimension(
                            new XlsbRecordSlice(
                                bytes,
                                recordOffset,
                                recordType,
                                payloadOffset,
                                recordSize),
                            out _,
                            out _);
                        sawDimension = true;
                        continue;
                    }
                    if (recordType == BrtBeginSheetData) {
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
                    if (recordType == BrtEndSheetData) {
                        if (!inSheetData) {
                            throw new InvalidDataException(
                                "The XLSB worksheet contains BrtEndSheetData without a matching BrtBeginSheetData record.");
                        }

                        if (previousCellColumn > lastColumn) {
                            lastColumn = previousCellColumn;
                        }
                        sawEndSheetData = true;
                        inSheetData = false;
                        currentRow = -1;
                        continue;
                    }
                    if (!inSheetData) {
                        if (recordType == BrtRowHdr || IsCellRecord(recordType)) {
                            throw new InvalidDataException(
                                $"The XLSB row or cell record at offset {recordOffset} appears outside BrtBeginSheetData/BrtEndSheetData.");
                        }

                        continue;
                    }
                    if (recordType == BrtRowHdr) {
                        if (previousCellColumn > lastColumn) {
                            lastColumn = previousCellColumn;
                        }
                        int nextRow = ValidateRowHeader(
                            new XlsbRecordSlice(
                                bytes,
                                recordOffset,
                                recordType,
                                payloadOffset,
                                recordSize),
                            currentRowSpanBounds,
                            out currentRowSpanCount);
                        if (currentRow >= 0 && nextRow <= currentRow) {
                            throw new InvalidDataException(
                                $"The XLSB worksheet contains non-increasing row index {nextRow} after header row {currentRow}.");
                        }

                        currentRow = nextRow;
                        previousCellColumn = -1;
                        rowPlanBuilder.BeginRow(nextRow, position);
                        continue;
                    }
                    if (!IsCellRecord(recordType)) {
                        continue;
                    }
                    if (currentRow < 0) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {recordOffset} appears before a row header.");
                    }

                    pendingCellBudget++;
                    if (pendingCellBudget == DiscoveryBudgetBatchSize) {
                        cellBudget.Consume(pendingCellBudget);
                        pendingCellBudget = 0;
                    }
                    if (!useCachedFormulaResult) {
                        EnsureFormulaModeSupported(recordType, useCachedFormulaResult);
                    }
                    int column;
                    if (recordSize >= sizeof(int) + sizeof(uint)
                        && recordType is >= BrtCellBlank and <= BrtCellIsst) {
                        ref byte payload = ref bytes[payloadOffset];
                        column = Unsafe.ReadUnaligned<int>(ref payload);
                        uint styleIndex = Unsafe.ReadUnaligned<uint>(
                            ref Unsafe.Add(ref payload, sizeof(int))) & 0x00FFFFFFU;
                        if (styleIndex >= styleCount) {
                            throw new InvalidDataException(
                                $"The XLSB cell record at offset {recordOffset} refers to missing cell format " +
                                $"{styleIndex}; the styles part exposes {styleCount} format(s).");
                        }

                        int valueBytes = recordSize - sizeof(int) - sizeof(uint);
                        bool validFixedPayload = recordType switch {
                            BrtCellBlank => true,
                            BrtCellRk => valueBytes >= sizeof(uint),
                            BrtCellError or BrtCellBool => valueBytes >= sizeof(byte),
                            BrtCellReal => valueBytes >= sizeof(double),
                            BrtCellIsst => valueBytes >= sizeof(uint),
                            _ => false
                        };
                        if (validFixedPayload && recordType == BrtCellIsst) {
                            uint sharedStringIndex = Unsafe.ReadUnaligned<uint>(
                                ref Unsafe.Add(ref payload, sizeof(int) + sizeof(uint)));
                            if (sharedStringIndex >= sharedStringCount) {
                                throw new InvalidDataException(
                                    $"The XLSB cell record at offset {recordOffset} refers to missing shared string " +
                                    $"{sharedStringIndex}; the shared-string part exposes {sharedStringCount} item(s).");
                            }
                        }
                        if (!validFixedPayload) {
                            column = ValidateCellPayloadStructure(
                                new XlsbRecordSlice(
                                    bytes,
                                    recordOffset,
                                    recordType,
                                    payloadOffset,
                                    recordSize),
                                styleCount,
                                sharedStringCount,
                                limits.MaxStringCharacters);
                        }
                    } else {
                        column = ValidateCellPayloadStructure(
                            new XlsbRecordSlice(
                                bytes,
                                recordOffset,
                                recordType,
                                payloadOffset,
                                recordSize),
                            styleCount,
                            sharedStringCount,
                            limits.MaxStringCharacters);
                    }
                    if (column < 0 || column >= A1.MaxColumns) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {recordOffset} contains invalid column index {column}.");
                    }
                    bool firstCellInRow = previousCellColumn < 0;
                    if (!firstCellInRow && column <= previousCellColumn) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {recordOffset} is duplicated or out of order within its row.");
                    }
                    previousCellColumn = column;
                    rowPlanBuilder.ObserveCell(
                        recordType,
                        payloadOffset,
                        recordSize,
                        column);
                    bool covered = currentRowSpanCount == 1
                        ? currentRowSpanBounds[0] <= column && column <= currentRowSpanBounds[1]
                        : IsCoveredByRowSpan(currentRowSpanBounds, currentRowSpanCount, column);
                    if (!covered) {
                        throw new InvalidDataException(
                            $"The XLSB cell record at offset {recordOffset} for column {column} is not covered by its BrtRowHdr column spans.");
                    }

                    if (firstCellInRow) {
                        if (firstDataRow < 0) {
                            firstDataRow = currentRow;
                        }
                        lastDataRow = currentRow;
                        if (column < firstColumn) {
                            firstColumn = column;
                        }
                    }
                }

                cancellationToken.ThrowIfCancellationRequested();
                recordBudget.Consume(pendingRecordBudget);
                cellBudget.Consume(pendingCellBudget);
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
                validatedRowPlan = rowPlanBuilder.Build();
            }
        }

        private static bool IsCoveredByRowSpan(int[] spanBounds, int spanCount, int column) {
            for (int index = 0; index < spanCount; index++) {
                int offset = index * 2;
                if (spanBounds[offset] <= column && column <= spanBounds[offset + 1]) {
                    return true;
                }
            }
            return false;
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
            if (record.Size >= sizeof(int) + sizeof(uint)) {
                byte[] bytes = record.Bytes;
                int position = record.PayloadOffset;
                int column = BinaryPrimitives.ReadInt32LittleEndian(
                    bytes.AsSpan(position, sizeof(int)));
                uint styleIndex = BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(position + sizeof(int), sizeof(uint))) & 0x00FFFFFFU;
                if (styleIndex >= styleCount) {
                    throw new InvalidDataException(
                        $"The XLSB cell record at offset {record.RecordOffset} refers to missing cell format " +
                        $"{styleIndex}; the styles part exposes {styleCount} format(s).");
                }

                int valuePosition = position + sizeof(int) + sizeof(uint);
                int valueBytes = record.Size - sizeof(int) - sizeof(uint);
                switch (record.Type) {
                    case BrtCellBlank:
                        return column;
                    case BrtCellRk:
                    case BrtCellIsst:
                        if (valueBytes >= sizeof(uint)) {
                            if (record.Type == BrtCellIsst) {
                                uint sharedStringIndex = BinaryPrimitives.ReadUInt32LittleEndian(
                                    bytes.AsSpan(valuePosition, sizeof(uint)));
                                if (sharedStringIndex >= sharedStringCount) {
                                    throw new InvalidDataException(
                                        $"The XLSB cell record at offset {record.RecordOffset} refers to missing shared string " +
                                        $"{sharedStringIndex}; the shared-string part exposes {sharedStringCount} item(s).");
                                }
                            }
                            return column;
                        }
                        break;
                    case BrtCellError:
                    case BrtCellBool:
                        if (valueBytes >= sizeof(byte)) {
                            return column;
                        }
                        break;
                    case BrtCellReal:
                        if (valueBytes >= sizeof(double)) {
                            return column;
                        }
                        break;
                }
            }

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
