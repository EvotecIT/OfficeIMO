using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private void StoreCellFast(XlsbRecordSlice record) {
            EnsureFormulaModeSupported(record.Type);
            var cursor = record.CreateCursor();
            int column = cursor.ReadInt32();
            uint styleIndex = cursor.ReadUInt32() & 0x00FFFFFFU;
            if (column < 0 || column >= A1.MaxColumns) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} contains invalid column index {column}.");
            }
            ValidateStyleIndex(styleIndex, record);

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
                    StoreNumber(ordinal, BiffRkNumberReader.ReadRkNumber(cursor.ReadUInt32()), styleIndex);
                    break;
                case BrtCellError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(cursor.ReadByte());
                    break;
                case BrtCellBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = cursor.ReadByte() != 0;
                    break;
                case BrtCellReal:
                    StoreNumber(ordinal, cursor.ReadDouble(), styleIndex);
                    break;
                case BrtCellSt:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    break;
                case BrtCellIsst: {
                    uint sharedStringIndex = cursor.ReadUInt32();
                    if (sharedStringIndex >= _sharedStrings.Count) {
                        throw new InvalidDataException(
                            $"The XLSB cell refers to missing shared string {sharedStringIndex}.");
                    }

                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = _sharedStrings[checked((int)sharedStringIndex)];
                    break;
                }
                case BrtCellRString:
                    cursor.ReadByte();
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    break;
                case BrtFmlaString:
                    _kinds[ordinal] = XlsbTabularValueKind.Text;
                    _strings[ordinal] = cursor.ReadWideString(_limits.MaxStringCharacters);
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaNum:
                    StoreNumber(ordinal, cursor.ReadDouble(), styleIndex);
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaBool:
                    _kinds[ordinal] = XlsbTabularValueKind.Boolean;
                    _booleans[ordinal] = cursor.ReadByte() != 0;
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                case BrtFmlaError:
                    _kinds[ordinal] = XlsbTabularValueKind.Error;
                    _strings[ordinal] = BiffErrorValue.ToText(cursor.ReadByte());
                    ValidateFormulaPayloadTail(record, ref cursor);
                    break;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell record type {record.Type}.");
            }
        }

        private void StoreNumber(int ordinal, double number, uint styleIndex) {
            bool isDate = _options.TreatDatesUsingNumberFormat
                && styleIndex < _dateStyles.Length
                && _dateStyles[styleIndex];
            _kinds[ordinal] = isDate ? XlsbTabularValueKind.Date : XlsbTabularValueKind.Number;
            _numbers[ordinal] = number;
        }

        private void ValidateStyleIndex(uint styleIndex, XlsbRecordSlice record) {
            if (styleIndex >= _dateStyles.Length) {
                throw new InvalidDataException(
                    $"The XLSB cell record at offset {record.RecordOffset} refers to missing cell format " +
                    $"{styleIndex}; the styles part exposes {_dateStyles.Length} format(s).");
            }
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
