using OfficeIMO.Excel.LegacyXls.Biff;
using System.Buffers.Binary;
using System.Runtime.CompilerServices;
using static OfficeIMO.Excel.LegacyXls.Read.LegacyXlsTabularWorkbook;

namespace OfficeIMO.Excel.LegacyXls.Read {
    internal sealed partial class LegacyXlsTabularDataReader {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void StoreBufferedCellRecord(
            byte[] bytes,
            int expectedRow,
            ushort type,
            int recordOffset,
            int payloadOffset,
            int length,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            switch ((BiffRecordType)type) {
                case BiffRecordType.LabelSst:
                case BiffRecordType.Number:
                case BiffRecordType.Rk:
                case BiffRecordType.BoolErr:
                    break;
                default:
                    StoreBufferedComplexCellRecord(bytes, expectedRow, type, recordOffset,
                        payloadOffset, length, ref pendingFormulaOrdinal, ref pendingFormulaStyle);
                    return;
            }
            int column = ReadBufferedCellColumn(bytes, expectedRow, type, recordOffset, payloadOffset, length);
            int ordinal = GetBufferedCellOrdinal(column);
            ushort style = (ushort)(bytes[payloadOffset + 4] | bytes[payloadOffset + 5] << 8);
            switch ((BiffRecordType)type) {
                case BiffRecordType.LabelSst:
                    if (length < 10) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = GetSharedString(BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(payloadOffset + 6, sizeof(uint))));
                    break;
                case BiffRecordType.Number:
                    if (length < 14) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    StoreNumber(ordinal, ReadBufferedDouble(bytes, payloadOffset + 6), style);
                    break;
                case BiffRecordType.Rk:
                    if (length < 10) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    StoreNumber(
                        ordinal,
                        BiffRkNumberReader.ReadRkNumber(BinaryPrimitives.ReadUInt32LittleEndian(
                            bytes.AsSpan(payloadOffset + 6, sizeof(uint)))),
                        style);
                    break;
                case BiffRecordType.BoolErr:
                    if (length < 8) throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
                    if (bytes[payloadOffset + 7] != 0) {
                        _kinds[ordinal] = ValueKind.Error;
                        _strings[ordinal] = BiffErrorValue.ToText(bytes[payloadOffset + 6]);
                    } else {
                        _kinds[ordinal] = ValueKind.Boolean;
                        _booleans[ordinal] = bytes[payloadOffset + 6] != 0;
                    }
                    break;
            }
        }

        private void StoreBufferedComplexCellRecord(
            byte[] bytes,
            int expectedRow,
            ushort type,
            int recordOffset,
            int payloadOffset,
            int length,
            ref int pendingFormulaOrdinal,
            ref ushort pendingFormulaStyle) {
            if (!IsCellRecordType(type)) return;
            int column = ReadBufferedCellColumn(bytes, expectedRow, type, recordOffset, payloadOffset, length);
            if (type == (ushort)BiffRecordType.MulRk) {
                StoreBufferedMulRk(bytes, payloadOffset, length, column);
                return;
            }
            if (type == (ushort)BiffRecordType.MulBlank) return;

            int ordinal = GetBufferedCellOrdinal(column);
            ushort style = (ushort)(bytes[payloadOffset + 4] | bytes[payloadOffset + 5] << 8);
            switch ((BiffRecordType)type) {
                case BiffRecordType.Blank:
                    _kinds[ordinal] = ValueKind.Empty;
                    break;
                case BiffRecordType.Label: {
                    var record = new RecordSlice(type, recordOffset, payloadOffset, length);
                    _kinds[ordinal] = ValueKind.Text;
                    _strings[ordinal] = ReadLabel(record);
                    break;
                }
                case BiffRecordType.Formula:
                    StoreFormula(
                        new RecordSlice(type, recordOffset, payloadOffset, length),
                        ordinal,
                        style,
                        ref pendingFormulaOrdinal,
                        ref pendingFormulaStyle);
                    break;
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int ReadBufferedCellColumn(
            byte[] bytes, int expectedRow, ushort type, int recordOffset, int payloadOffset, int length) {
            if (length < 6) {
                throw TruncatedCell(new RecordSlice(type, recordOffset, payloadOffset, length));
            }
            uint rowAndColumn = BinaryPrimitives.ReadUInt32LittleEndian(
                bytes.AsSpan(payloadOffset, sizeof(uint)));
            int row = (ushort)rowAndColumn;
            if (row != expectedRow) {
                throw new InvalidDataException(
                    $"The indexed XLS row span for row {expectedRow} contains a cell record for row {row}.");
            }
            return (ushort)(rowAndColumn >> 16);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int GetBufferedCellOrdinal(int column) {
            int ordinal = column - _firstColumn;
            if ((uint)ordinal >= (uint)FieldCount) {
                throw new InvalidDataException(
                    $"The XLS row contains column {column} outside its discovered schema.");
            }
            return ordinal;
        }

        private void StoreBufferedMulRk(
            byte[] bytes,
            int payloadOffset,
            int length,
            int firstColumn) {
            if (length < 6) {
                throw new InvalidDataException("A MULRK record has an invalid payload length.");
            }
            int lastColumn = bytes[payloadOffset + length - 2]
                | bytes[payloadOffset + length - 1] << 8;
            int count = checked(lastColumn - firstColumn + 1);
            int expectedLength = checked(4 + count * 6 + 2);
            if (length != expectedLength) {
                throw new InvalidDataException("A MULRK record has an invalid payload length.");
            }
            int ordinal = firstColumn - _firstColumn;
            if (ordinal < 0 || ordinal > FieldCount - count) {
                throw new InvalidDataException(
                    $"The XLS row contains columns {firstColumn}..{lastColumn} outside its discovered schema.");
            }
            int cellOffset = payloadOffset + 4;
            for (int index = 0; index < count; index++, ordinal++, cellOffset += 6) {
                StoreNumber(
                    ordinal,
                    BiffRkNumberReader.ReadRkNumber(BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(cellOffset + 2, sizeof(uint)))),
                    BinaryPrimitives.ReadUInt16LittleEndian(bytes.AsSpan(cellOffset, sizeof(ushort))));
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static double ReadBufferedDouble(byte[] bytes, int offset) =>
            BitConverter.Int64BitsToDouble(BinaryPrimitives.ReadInt64LittleEndian(
                bytes.AsSpan(offset, sizeof(long))));

    }
}
