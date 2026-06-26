using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffFormulaArrayConstantReader {
        internal static bool TryRead(byte[] formulaPayload, ref int extraOffset, out string? formulaText) {
            formulaText = null;
            try {
                if (extraOffset + 3 > formulaPayload.Length) {
                    return false;
                }

                int columnCount = checked(formulaPayload[extraOffset++] + 1);
                int rowCount = checked(BiffRecordReader.ReadUInt16(formulaPayload, extraOffset) + 1);
                extraOffset += 2;

                var rows = new string[rowCount];
                for (int row = 0; row < rowCount; row++) {
                    var values = new string[columnCount];
                    for (int column = 0; column < columnCount; column++) {
                        if (!TryReadValue(formulaPayload, ref extraOffset, out string? value)) {
                            return false;
                        }

                        values[column] = value!;
                    }

                    rows[row] = string.Join(",", values);
                }

                formulaText = "{" + string.Join(";", rows) + "}";
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        private static bool TryReadValue(byte[] formulaPayload, ref int offset, out string? value) {
            value = null;
            if (offset >= formulaPayload.Length) {
                return false;
            }

            byte kind = formulaPayload[offset];
            switch (kind) {
                case 0x00:
                    if (offset + 9 > formulaPayload.Length) return false;
                    value = string.Empty;
                    offset += 9;
                    return true;
                case 0x01:
                    if (offset + 9 > formulaPayload.Length) return false;
                    value = BiffRecordReader.ReadDouble(formulaPayload, offset + 1).ToString("G15", CultureInfo.InvariantCulture);
                    offset += 9;
                    return true;
                case 0x02:
                    offset++;
                    value = QuoteFormulaStringLiteral(BiffStringReader.ReadUnicodeString(formulaPayload, ref offset));
                    return true;
                case 0x04:
                    if (offset + 9 > formulaPayload.Length) return false;
                    value = formulaPayload[offset + 1] == 0 ? "FALSE" : "TRUE";
                    offset += 9;
                    return true;
                case 0x10:
                    if (offset + 9 > formulaPayload.Length) return false;
                    value = BiffErrorValue.ToText(formulaPayload[offset + 1]);
                    offset += 9;
                    return true;
                default:
                    return false;
            }
        }

        private static string QuoteFormulaStringLiteral(string value) {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
