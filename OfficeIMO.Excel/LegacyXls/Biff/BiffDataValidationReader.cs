using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffDataValidationReader {
        private const uint AllowBlankFlag = 0x00000100;
        private const uint ShowInputMessageFlag = 0x00040000;
        private const uint ShowErrorMessageFlag = 0x00080000;

        internal static bool TryRead(byte[] payload, out LegacyXlsDataValidation? validation) {
            return TryRead(payload, Array.Empty<BiffExternSheetReference>(), Array.Empty<LegacyXlsExternalReference>(), Array.Empty<string>(), Array.Empty<string?>(), out validation);
        }

        internal static bool TryRead(byte[] payload, IReadOnlyList<string?> definedNames, out LegacyXlsDataValidation? validation) {
            return TryRead(payload, Array.Empty<BiffExternSheetReference>(), Array.Empty<LegacyXlsExternalReference>(), Array.Empty<string>(), definedNames, out validation);
        }

        internal static bool TryRead(
            byte[] payload,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out LegacyXlsDataValidation? validation) {
            return TryRead(payload, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, definedNames, out validation);
        }

        internal static bool TryRead(
            byte[] payload,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out LegacyXlsDataValidation? validation) {
            validation = null;
            if (payload.Length < 4) {
                return false;
            }

            int offset = 0;
            uint flags = BiffRecordReader.ReadUInt32(payload, offset);
            offset += 4;

            uint validationType = flags & 0x0000000f;
            uint operatorType = (flags >> 20) & 0x0000000f;
            if (!IsSupportedValidationType(validationType) || (validationType != 0x03 && validationType != 0x07 && operatorType > 0x07)) {
                return false;
            }

            LegacyXlsDataValidationType modelValidationType = ToValidationType(validationType);
            LegacyXlsDataValidationOperator comparisonOperator = ToOperator(operatorType);
            try {
                string promptTitle = BiffStringReader.ReadUnicodeString(payload, ref offset);
                string errorTitle = BiffStringReader.ReadUnicodeString(payload, ref offset);
                string prompt = BiffStringReader.ReadUnicodeString(payload, ref offset);
                string error = BiffStringReader.ReadUnicodeString(payload, ref offset);

                if (!TryReadFormula(payload, ref offset, externSheets, externalReferences, sheetNames, definedNames, out string? formula1) || string.IsNullOrWhiteSpace(formula1)) {
                    return false;
                }

                if (!TryReadFormula(payload, ref offset, externSheets, externalReferences, sheetNames, definedNames, out string? formula2)) {
                    return false;
                }

                bool requiresSecondFormula = modelValidationType != LegacyXlsDataValidationType.List
                    && modelValidationType != LegacyXlsDataValidationType.Custom
                    && (comparisonOperator == LegacyXlsDataValidationOperator.Between
                        || comparisonOperator == LegacyXlsDataValidationOperator.NotBetween);
                if (requiresSecondFormula && string.IsNullOrWhiteSpace(formula2)) {
                    return false;
                }

                if (!TryGetListSource(modelValidationType, formula1!, out IReadOnlyList<string>? listItems, out string? listSourceRange, out string? listSourceName, out string? listSourceSheetName)
                    || !IsSupportedFormula(modelValidationType, formula1!, listItems, listSourceRange, listSourceName)
                    || (modelValidationType != LegacyXlsDataValidationType.List
                        && modelValidationType != LegacyXlsDataValidationType.Custom
                        && formula2 != null
                        && !IsSupportedFormula(modelValidationType, formula2, null, null, null))) {
                    return false;
                }

                if (!TryReadRanges(payload, ref offset, out IReadOnlyList<string>? ranges) || ranges.Count == 0) {
                    return false;
                }

                validation = new LegacyXlsDataValidation(
                    modelValidationType,
                    comparisonOperator,
                    formula1!,
                    formula2,
                    (flags & AllowBlankFlag) != 0,
                    (flags & ShowInputMessageFlag) != 0,
                    (flags & ShowErrorMessageFlag) != 0,
                    EmptyToNull(promptTitle),
                    EmptyToNull(prompt),
                    EmptyToNull(errorTitle),
                    EmptyToNull(error),
                    ranges,
                    listItems,
                    listSourceRange,
                    listSourceName,
                    listSourceSheetName);
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        private static bool IsSupportedValidationType(uint validationType) {
            return validationType >= 0x01 && validationType <= 0x07;
        }

        private static LegacyXlsDataValidationType ToValidationType(uint validationType) {
            return validationType switch {
                0x01 => LegacyXlsDataValidationType.WholeNumber,
                0x02 => LegacyXlsDataValidationType.Decimal,
                0x03 => LegacyXlsDataValidationType.List,
                0x04 => LegacyXlsDataValidationType.Date,
                0x05 => LegacyXlsDataValidationType.Time,
                0x06 => LegacyXlsDataValidationType.TextLength,
                _ => LegacyXlsDataValidationType.Custom
            };
        }

        private static LegacyXlsDataValidationOperator ToOperator(uint operatorType) {
            return operatorType switch {
                0x00 => LegacyXlsDataValidationOperator.Between,
                0x01 => LegacyXlsDataValidationOperator.NotBetween,
                0x02 => LegacyXlsDataValidationOperator.Equal,
                0x03 => LegacyXlsDataValidationOperator.NotEqual,
                0x04 => LegacyXlsDataValidationOperator.GreaterThan,
                0x05 => LegacyXlsDataValidationOperator.LessThan,
                0x06 => LegacyXlsDataValidationOperator.GreaterThanOrEqual,
                _ => LegacyXlsDataValidationOperator.LessThanOrEqual
            };
        }

        private static bool TryReadFormula(
            byte[] payload,
            ref int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formula) {
            formula = null;
            if (offset + 4 > payload.Length) {
                return false;
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(payload, offset);
            offset += 2;
            offset += 2;
            if (offset + expressionLength > payload.Length) {
                return false;
            }

            if (expressionLength == 0) {
                return true;
            }

            byte[] normalizedFormula = new byte[checked(expressionLength + 2)];
            normalizedFormula[0] = (byte)(expressionLength & 0x00ff);
            normalizedFormula[1] = (byte)(expressionLength >> 8);
            Buffer.BlockCopy(payload, offset, normalizedFormula, 2, expressionLength);
            offset += expressionLength;
            return BiffFormulaTextReader.TryRead(
                normalizedFormula,
                0,
                formulaRow: 0,
                formulaColumn: 0,
                externSheets,
                externalReferences,
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
            if (count == 0 || count > 432) {
                return false;
            }

            int expectedLength = checked(count * 8);
            if (offset + expectedLength > payload.Length) {
                return false;
            }

            var parsedRanges = new List<string>(count);
            for (int i = 0; i < count; i++) {
                ushort firstRow = BiffRecordReader.ReadUInt16(payload, offset);
                ushort lastRow = BiffRecordReader.ReadUInt16(payload, offset + 2);
                ushort firstColumn = BiffRecordReader.ReadUInt16(payload, offset + 4);
                ushort lastColumn = BiffRecordReader.ReadUInt16(payload, offset + 6);
                offset += 8;

                if (lastRow < firstRow || lastColumn < firstColumn || firstColumn > 0x00ff || lastColumn > 0x00ff) {
                    return false;
                }

                string start = A1.CellReference(firstRow + 1, firstColumn + 1);
                string end = A1.CellReference(lastRow + 1, lastColumn + 1);
                parsedRanges.Add(start == end ? start : start + ":" + end);
            }

            ranges = parsedRanges;
            return true;
        }

        private static bool IsSupportedFormula(
            LegacyXlsDataValidationType validationType,
            string formula,
            IReadOnlyList<string>? listItems,
            string? listSourceRange,
            string? listSourceName) {
            return validationType switch {
                LegacyXlsDataValidationType.WholeNumber => int.TryParse(formula, NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
                LegacyXlsDataValidationType.Decimal => double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out _),
                LegacyXlsDataValidationType.List => listItems?.Count > 0 || !string.IsNullOrWhiteSpace(listSourceRange) || !string.IsNullOrWhiteSpace(listSourceName),
                LegacyXlsDataValidationType.Date => double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out _),
                LegacyXlsDataValidationType.Time => double.TryParse(formula, NumberStyles.Float, CultureInfo.InvariantCulture, out _),
                LegacyXlsDataValidationType.TextLength => int.TryParse(formula, NumberStyles.Integer, CultureInfo.InvariantCulture, out _),
                LegacyXlsDataValidationType.Custom => !string.IsNullOrWhiteSpace(formula),
                _ => false
            };
        }

        private static bool TryGetListSource(
            LegacyXlsDataValidationType validationType,
            string formula,
            out IReadOnlyList<string>? listItems,
            out string? listSourceRange,
            out string? listSourceName,
            out string? listSourceSheetName) {
            listItems = null;
            listSourceRange = null;
            listSourceName = null;
            listSourceSheetName = null;
            if (validationType != LegacyXlsDataValidationType.List) {
                return true;
            }

            if (TryParseInlineListFormula(formula, out IReadOnlyList<string>? parsedItems) && parsedItems.Count > 0) {
                listItems = parsedItems;
                return true;
            }

            if (TryParseSheetQualifiedListSourceRange(formula, out listSourceSheetName, out listSourceRange)) {
                return true;
            }

            if (TryParseListSourceRange(formula, out listSourceRange)) {
                return true;
            }

            return TryParseListSourceName(formula, out listSourceName);
        }

        private static bool TryParseInlineListFormula(string formula, out IReadOnlyList<string> items) {
            items = Array.Empty<string>();
            if (formula.Length < 2 || formula[0] != '"' || formula[formula.Length - 1] != '"') {
                return false;
            }

            string inner = formula.Substring(1, formula.Length - 2).Replace("\"\"", "\"");
            if (inner.Length == 0) {
                return false;
            }

            items = inner.Split(',')
                .Select(item => RemoveInvalidXmlCharacters(item.Trim()))
                .Where(item => item.Length > 0)
                .ToArray();
            return items.Count > 0;
        }

        private static bool TryParseSheetQualifiedListSourceRange(string formula, out string? sourceSheetName, out string? sourceRange) {
            sourceSheetName = null;
            sourceRange = null;
            string normalizedFormula = formula.Trim();
            if (normalizedFormula.StartsWith("=", StringComparison.Ordinal)) {
                normalizedFormula = normalizedFormula.Substring(1).Trim();
            }

            if (!SheetNameLookup.TryParseSheetQualifiedReference(normalizedFormula, out string parsedSheetName, out string parsedReference, allowExternalWorkbookReferences: false)) {
                return false;
            }

            if (!TryParseListSourceRange(parsedReference, out string? parsedRange)) {
                return false;
            }

            sourceSheetName = parsedSheetName;
            sourceRange = parsedRange;
            return true;
        }

        private static bool TryParseListSourceRange(string formula, out string? sourceRange) {
            sourceRange = null;
            string normalized = formula.Trim().Replace("$", string.Empty);
            if (normalized.Length == 0 || normalized.IndexOfAny(new[] { '!', ',', ' ', '(', ')', '+', '-', '*', '/', '&' }) >= 0) {
                return false;
            }

            if (normalized.IndexOf(':') >= 0) {
                if (!A1.TryParseRange(normalized, out int startRow, out int startColumn, out int endRow, out int endColumn)) {
                    return false;
                }

                string start = A1.CellReference(startRow, startColumn);
                string end = A1.CellReference(endRow, endColumn);
                sourceRange = start == end ? start : start + ":" + end;
                return true;
            }

            if (!A1.TryParseCellReferenceFast(normalized, out int row, out int column)) {
                return false;
            }

            sourceRange = A1.CellReference(row, column);
            return true;
        }

        private static bool TryParseListSourceName(string formula, out string? sourceName) {
            sourceName = null;
            string normalized = formula.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal)) {
                normalized = normalized.Substring(1).Trim();
            }

            if (normalized.Length == 0 || normalized.IndexOfAny(new[] { '!', ':', ',', ' ', '\t', '\r', '\n', '(', ')', '+', '-', '*', '/', '&', '"', '\'' }) >= 0) {
                return false;
            }

            char first = normalized[0];
            if (!char.IsLetter(first) && first != '_' && first != '\\') {
                return false;
            }

            for (int i = 1; i < normalized.Length; i++) {
                char current = normalized[i];
                if (!char.IsLetterOrDigit(current) && current != '_' && current != '.' && current != '\\') {
                    return false;
                }
            }

            if (A1.TryParseCellReferenceFast(normalized, out _, out _)) {
                return false;
            }

            sourceName = normalized;
            return true;
        }

        private static string? EmptyToNull(string value) {
            string normalized = RemoveInvalidXmlCharacters(value);
            return normalized.Length == 0 ? null : normalized;
        }

        private static string RemoveInvalidXmlCharacters(string value) {
            StringBuilder? builder = null;
            for (int i = 0; i < value.Length; i++) {
                char current = value[i];
                if (IsXmlCharacter(current)) {
                    builder?.Append(current);
                    continue;
                }

                builder ??= new StringBuilder(value.Length).Append(value, 0, i);
            }

            return builder == null ? value : builder.ToString();
        }

        private static bool IsXmlCharacter(char value) {
            return value == '\t'
                || value == '\n'
                || value == '\r'
                || (value >= ' ' && value <= '\ud7ff')
                || (value >= '\ue000' && value <= '\ufffd');
        }
    }
}
