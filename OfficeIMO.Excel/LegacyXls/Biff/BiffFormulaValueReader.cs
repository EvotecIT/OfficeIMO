using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal readonly struct BiffFormulaValue {
        internal BiffFormulaValue(LegacyXlsCellValueKind kind, object? value, bool expectsStringRecord) {
            Kind = kind;
            Value = value;
            ExpectsStringRecord = expectsStringRecord;
        }

        internal LegacyXlsCellValueKind Kind { get; }

        internal object? Value { get; }

        internal bool ExpectsStringRecord { get; }
    }

    internal static class BiffFormulaValueReader {
        internal static BiffFormulaValue Read(byte[] payload, int offset) {
            if (offset + 8 > payload.Length) {
                throw new InvalidDataException("The FormulaValue structure ended early.");
            }

            ushort fExprO = BiffRecordReader.ReadUInt16(payload, offset + 6);
            if (fExprO != 0xffff) {
                return new BiffFormulaValue(
                    LegacyXlsCellValueKind.Number,
                    BiffRecordReader.ReadDouble(payload, offset),
                    expectsStringRecord: false);
            }

            byte valueType = payload[offset];
            switch (valueType) {
                case 0x00:
                    return new BiffFormulaValue(LegacyXlsCellValueKind.Text, null, expectsStringRecord: true);
                case 0x01:
                    return new BiffFormulaValue(LegacyXlsCellValueKind.Boolean, payload[offset + 2] != 0, expectsStringRecord: false);
                case 0x02:
                    return new BiffFormulaValue(LegacyXlsCellValueKind.Error, BiffErrorValue.ToText(payload[offset + 2]), expectsStringRecord: false);
                case 0x03:
                    return new BiffFormulaValue(LegacyXlsCellValueKind.Text, string.Empty, expectsStringRecord: false);
                default:
                    throw new InvalidDataException($"FormulaValue type 0x{valueType:X2} is not supported.");
            }
        }
    }
}
