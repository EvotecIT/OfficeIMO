using OfficeIMO.Excel.Xlsb.Biff12;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        /// <summary>
        /// Proves one cell record's payload against the worksheet's styles, shared strings and
        /// string limits, and returns its column. The common fixed-size shapes are checked inline.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int ValidateCellRecord(
            byte[] bytes,
            int recordType,
            int recordOffset,
            int payloadOffset,
            int recordSize) {
            if (!_options.UseCachedFormulaResult) {
                EnsureFormulaModeSupported(recordType, useCachedFormulaResult: false);
            }

            int styleCount = _dateStyles.Length;
            int column;
            if (recordSize >= sizeof(int) + sizeof(uint)
                && recordType is >= BrtCellBlank and <= BrtCellIsst) {
                ref byte payload = ref bytes[payloadOffset];
                column = Unsafe.ReadUnaligned<int>(ref payload);
                uint styleIndex = Unsafe.ReadUnaligned<uint>(
                    ref Unsafe.Add(ref payload, sizeof(int))) & 0x00FFFFFFU;
                if (styleIndex >= styleCount) {
                    ThrowMissingCellFormat(recordOffset, styleIndex, styleCount);
                }

                int valueBytes = recordSize - sizeof(int) - sizeof(uint);
                bool validFixedPayload = recordType switch {
                    BrtCellBlank => true,
                    BrtCellRk => valueBytes >= sizeof(uint),
                    BrtCellError or BrtCellBool => valueBytes >= sizeof(byte),
                    BrtCellReal => valueBytes >= sizeof(double),
                    BrtCellIsst => valueBytes >= sizeof(uint),
                    // A fitting inline string; any other shape keeps the cursor's exact errors.
                    BrtCellSt => valueBytes >= sizeof(uint)
                        && Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref payload, sizeof(int) + sizeof(uint))) is uint characters
                        && characters <= (long)_limits.MaxStringCharacters
                        && characters <= (uint)(valueBytes - sizeof(uint)) / sizeof(char),
                    _ => false
                };
                if (validFixedPayload && recordType == BrtCellIsst) {
                    uint sharedStringIndex = Unsafe.ReadUnaligned<uint>(
                        ref Unsafe.Add(ref payload, sizeof(int) + sizeof(uint)));
                    if (sharedStringIndex >= _sharedStrings.Count) {
                        ThrowMissingSharedString(recordOffset, sharedStringIndex, _sharedStrings.Count);
                    }
                }
                if (!validFixedPayload) {
                    column = ValidateCellPayloadStructure(
                        new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize),
                        styleCount,
                        _sharedStrings.Count,
                        _limits.MaxStringCharacters);
                }
            } else {
                column = ValidateCellPayloadStructure(
                    new XlsbRecordSlice(bytes, recordOffset, recordType, payloadOffset, recordSize),
                    styleCount,
                    _sharedStrings.Count,
                    _limits.MaxStringCharacters);
            }
            if (column < 0 || column >= ExcelA1.MaxColumns) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {recordOffset} contains invalid column index {column}.");
            }

            return column;
        }

        /// <summary>
        /// Requires strictly increasing cell columns within a row, each covered by the row
        /// header's declared column spans.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ValidateCellPlacement(int column, int recordOffset, ref int previousCellColumn) {
            if (previousCellColumn >= 0 && column <= previousCellColumn) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {recordOffset} is duplicated or out of order within its row.");
            }

            previousCellColumn = column;
            bool covered = _rowSpanCount == 1
                ? _rowSpanBounds[0] <= column && column <= _rowSpanBounds[1]
                : IsCoveredByRowSpan(_rowSpanBounds, _rowSpanCount, column);
            if (!covered) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {recordOffset} for column {column} is not covered by its BrtRowHdr column spans.");
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowMissingCellFormat(int recordOffset, uint styleIndex, int styleCount) =>
            throw new InvalidDataException(
                $"The XLSB cell record at offset {recordOffset} refers to missing cell format " +
                $"{styleIndex}; the styles part exposes {styleCount} format(s).");

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ThrowMissingSharedString(int recordOffset, uint sharedStringIndex, int sharedStringCount) =>
            throw new InvalidDataException(
                $"The XLSB cell record at offset {recordOffset} refers to missing shared string " +
                $"{sharedStringIndex}; the shared-string part exposes {sharedStringCount} item(s).");

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
                || lastRow >= ExcelA1.MaxRows
                || firstColumnValue > lastColumnValue
                || lastColumnValue >= ExcelA1.MaxColumns) {
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
                        cursor.ValidateWideString(maxStringCharacters);
                        break;
                    case BrtCellRString:
                        cursor.ReadByte();
                        cursor.ValidateWideString(maxStringCharacters);
                        break;
                    case BrtFmlaString:
                        cursor.ValidateWideString(maxStringCharacters);
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
