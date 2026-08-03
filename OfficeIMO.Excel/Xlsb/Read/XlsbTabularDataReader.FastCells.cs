using OfficeIMO.Excel.LegacyXls.Biff;
using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private void StoreCellFast(XlsbRecordSlice record) {
            byte[] bytes = record.Bytes;
            int position = record.PayloadOffset;
            int column = BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(position, sizeof(int)));
            uint styleIndex = BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(position + sizeof(int), sizeof(uint)))
                & 0x00FFFFFFU;
            position += sizeof(int) + sizeof(uint);

            int ordinal = column - _firstColumn;
            if (ordinal < 0 || ordinal >= FieldCount) {
                throw new InvalidDataException(
                    $"The XLSB row contains column {column} outside the schema established by its header or worksheet dimension.");
            }

            switch (record.Type) {
                case BrtCellBlank:
                    _kinds[ordinal] = XlsbTabularValueKind.Empty;
                    break;
                case BrtCellRk:
                    StoreNumber(
                        ordinal,
                        BiffRkNumberReader.ReadRkNumber(
                            BinaryPrimitives.ReadUInt32LittleEndian(bytes.AsSpan(position, sizeof(uint)))),
                        styleIndex);
                    break;
                case BrtCellError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(bytes[position]);
                    break;
                case BrtCellBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = bytes[position] != 0;
                    break;
                case BrtCellReal:
                    StoreNumber(
                        ordinal,
                        BitConverter.Int64BitsToDouble(
                            BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(position, sizeof(long)))),
                        styleIndex);
                    break;
                case BrtCellSt:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = ReadValidatedWideString(bytes, position);
                    break;
                case BrtCellIsst: {
                    uint sharedStringIndex = BinaryPrimitives.ReadUInt32LittleEndian(
                        bytes.AsSpan(position, sizeof(uint)));
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = _sharedStrings[checked((int)sharedStringIndex)];
                    break;
                }
                case BrtCellRString:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = ReadValidatedWideString(bytes, position + 1);
                    break;
                case BrtFmlaString:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = ReadValidatedWideString(bytes, position);
                    break;
                case BrtFmlaNum:
                    StoreNumber(
                        ordinal,
                        BitConverter.Int64BitsToDouble(
                            BinaryPrimitives.ReadInt64LittleEndian(bytes.AsSpan(position, sizeof(long)))),
                        styleIndex);
                    break;
                case BrtFmlaBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = bytes[position] != 0;
                    break;
                case BrtFmlaError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(bytes[position]);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }
        }

        private static string ReadValidatedWideString(byte[] bytes, int position) {
            int characterCount = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(
                bytes.AsSpan(position, sizeof(uint))));
            return Encoding.Unicode.GetString(
                bytes,
                position + sizeof(uint),
                checked(characterCount * sizeof(char)));
        }

        private void StoreNumber(int ordinal, double number, uint styleIndex) {
            bool isDate = _options.TreatDatesUsingNumberFormat
                && styleIndex < _dateStyles.Length
                && _dateStyles[styleIndex];
            _kinds[ordinal] = isDate ? XlsbTabularValueKind.Date : XlsbTabularValueKind.Number;
            _numbers[ordinal] = number;
        }

        private static void ValidateFormulaPayloadTail(
            XlsbRecordSlice record,
            ref XlsbSliceCursor cursor) {
            const int mandatoryHeaderBytes = sizeof(ushort) + sizeof(uint);
            if (cursor.Remaining < mandatoryHeaderBytes) {
                throw new InvalidDataException(
                    $"The XLSB formula record at offset {record.RecordOffset} ended before its flags and token-byte count.");
            }

            cursor.ReadUInt16(); // grbit flags
            uint tokenCount = cursor.ReadUInt32();
            if (tokenCount > cursor.Remaining) {
                throw new InvalidDataException(
                    $"The XLSB formula record at offset {record.RecordOffset} declares {tokenCount} token bytes but only {cursor.Remaining} remain.");
            }

            cursor.Skip(checked((int)tokenCount));
        }
    }
}
