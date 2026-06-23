using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffConditionalFormattingReader {
        internal static bool TryReadHeader(byte[] payload, out ushort ruleCount, out IReadOnlyList<string> ranges) {
            ruleCount = 0;
            ranges = Array.Empty<string>();
            if (payload.Length < 14) {
                return false;
            }

            ushort ccf = BiffRecordReader.ReadUInt16(payload, 0);
            if (ccf == 0 || ccf > 1024) {
                return false;
            }

            if (!TryReadCellRange(payload, 4, out string? enclosingRange)) {
                return false;
            }

            int sqrefOffset = 12;
            if (!TryReadRanges(payload, ref sqrefOffset, out ranges) || ranges.Count == 0) {
                ranges = new[] { enclosingRange! };
            }

            ruleCount = ccf;
            return true;
        }

        internal static bool TryReadRule(
            byte[] payload,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            IReadOnlyList<string> ranges,
            out LegacyXlsConditionalFormatting? conditionalFormatting) {
            conditionalFormatting = null;
            if (payload.Length < 6 || ranges.Count == 0) {
                return false;
            }

            byte conditionType = payload[0];
            byte comparisonOperator = payload[1];
            ushort formula1Length = BiffRecordReader.ReadUInt16(payload, 2);
            ushort formula2Length = BiffRecordReader.ReadUInt16(payload, 4);
            int formulaStart = payload.Length - formula1Length - formula2Length;
            if (formulaStart < 6 || formulaStart > payload.Length) {
                return false;
            }

            if (!TryReadFormula(payload, formulaStart, formula1Length, externSheets, sheetNames, definedNames, out string? formula1)) {
                return false;
            }

            if (!TryReadFormula(payload, formulaStart + formula1Length, formula2Length, externSheets, sheetNames, definedNames, out string? formula2)) {
                return false;
            }

            if (conditionType == 0x01) {
                if (!TryMapOperator(comparisonOperator, out LegacyXlsConditionalFormattingOperator @operator)
                    || string.IsNullOrWhiteSpace(formula1)) {
                    return false;
                }

                bool requiresSecondFormula = @operator == LegacyXlsConditionalFormattingOperator.Between
                    || @operator == LegacyXlsConditionalFormattingOperator.NotBetween;
                if (requiresSecondFormula && string.IsNullOrWhiteSpace(formula2)) {
                    return false;
                }

                conditionalFormatting = new LegacyXlsConditionalFormatting(
                    LegacyXlsConditionalFormattingType.CellIs,
                    @operator,
                    formula1!,
                    string.IsNullOrWhiteSpace(formula2) ? null : formula2,
                    ranges);
                return true;
            }

            if (conditionType == 0x02) {
                if (string.IsNullOrWhiteSpace(formula1) || !string.IsNullOrWhiteSpace(formula2)) {
                    return false;
                }

                conditionalFormatting = new LegacyXlsConditionalFormatting(
                    LegacyXlsConditionalFormattingType.Formula,
                    null,
                    formula1!,
                    null,
                    ranges);
                return true;
            }

            return false;
        }

        private static bool TryReadFormula(
            byte[] payload,
            int offset,
            ushort expressionLength,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formula) {
            formula = null;
            if (expressionLength == 0) {
                return true;
            }

            if (offset < 0 || offset + expressionLength > payload.Length) {
                return false;
            }

            byte[] normalizedFormula = new byte[checked(expressionLength + 2)];
            normalizedFormula[0] = (byte)(expressionLength & 0x00ff);
            normalizedFormula[1] = (byte)(expressionLength >> 8);
            Buffer.BlockCopy(payload, offset, normalizedFormula, 2, expressionLength);
            return BiffFormulaTextReader.TryRead(
                normalizedFormula,
                0,
                formulaRow: 0,
                formulaColumn: 0,
                externSheets,
                sheetNames,
                definedNames,
                out formula);
        }

        private static bool TryReadRanges(byte[] payload, ref int offset, out IReadOnlyList<string> ranges) {
            ranges = Array.Empty<string>();
            if (offset + 2 > payload.Length) {
                return false;
            }

            ushort count = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            if (count == 0 || count > 8192) {
                return false;
            }

            int expectedLength = checked(count * 8);
            if (offset + expectedLength > payload.Length) {
                return false;
            }

            var parsedRanges = new List<string>(count);
            for (int i = 0; i < count; i++) {
                if (!TryReadCellRange(payload, offset, out string? range)) {
                    return false;
                }

                parsedRanges.Add(range!);
                offset += 8;
            }

            ranges = parsedRanges;
            return true;
        }

        private static bool TryReadCellRange(byte[] payload, int offset, out string? range) {
            range = null;
            if (offset + 8 > payload.Length) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
            if (lastRow < firstRow || lastColumn < firstColumn || firstColumn > 0x00ff || lastColumn > 0x00ff) {
                return false;
            }

            string start = A1.CellReference(firstRow + 1, firstColumn + 1);
            string end = A1.CellReference(lastRow + 1, lastColumn + 1);
            range = start == end ? start : start + ":" + end;
            return true;
        }

        private static bool TryMapOperator(byte value, out LegacyXlsConditionalFormattingOperator @operator) {
            switch (value) {
                case 0x01:
                    @operator = LegacyXlsConditionalFormattingOperator.Between;
                    return true;
                case 0x02:
                    @operator = LegacyXlsConditionalFormattingOperator.NotBetween;
                    return true;
                case 0x03:
                    @operator = LegacyXlsConditionalFormattingOperator.Equal;
                    return true;
                case 0x04:
                    @operator = LegacyXlsConditionalFormattingOperator.NotEqual;
                    return true;
                case 0x05:
                    @operator = LegacyXlsConditionalFormattingOperator.GreaterThan;
                    return true;
                case 0x06:
                    @operator = LegacyXlsConditionalFormattingOperator.LessThan;
                    return true;
                case 0x07:
                    @operator = LegacyXlsConditionalFormattingOperator.GreaterThanOrEqual;
                    return true;
                case 0x08:
                    @operator = LegacyXlsConditionalFormattingOperator.LessThanOrEqual;
                    return true;
                default:
                    @operator = LegacyXlsConditionalFormattingOperator.Equal;
                    return false;
            }
        }
    }
}
