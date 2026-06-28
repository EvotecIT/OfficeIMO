using System.Globalization;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Decodes the first supported NameParsedFormula shapes used by BIFF defined names.
    /// </summary>
    internal static class BiffNameFormulaReader {
        internal static bool TryReadReference(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryReadReference(formulaBytes, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out reference);
        }

        internal static bool TryReadReference(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryReadReference(formulaBytes, externSheets, externalReferences, sheetNames, out reference, out _);
        }

        internal static bool TryReadReference(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference,
            out BiffFormulaReadFailure? failure) {
            reference = null;
            failure = null;
            if (formulaBytes.Length == 0) {
                failure = BiffFormulaReadFailure.InvalidPayload("Defined-name formula token stream was empty.");
                return false;
            }

            int offset = 0;
            int tokenOffset = offset;
            byte token = formulaBytes[offset++];
            switch (token) {
                case 0x3a:
                case 0x5a:
                case 0x7a:
                    if (offset + 6 != formulaBytes.Length) {
                        failure = BiffFormulaReadFailure.InvalidPayload("Defined-name 3D reference formula has an unexpected length.", token, tokenOffset);
                        return false;
                    }

                    if (!BiffFormulaReferenceFormatter.TryRead3dReference(formulaBytes, offset, externSheets, externalReferences, sheetNames, out reference)) {
                        failure = BiffFormulaReadFailure.Reference("FormulaDefinedName3dReference", "Defined-name 3D reference could not be resolved.", token, tokenOffset);
                        return false;
                    }

                    return true;
                case 0x3b:
                case 0x5b:
                case 0x7b:
                    if (offset + 10 != formulaBytes.Length) {
                        failure = BiffFormulaReadFailure.InvalidPayload("Defined-name 3D area formula has an unexpected length.", token, tokenOffset);
                        return false;
                    }

                    if (!BiffFormulaReferenceFormatter.TryRead3dArea(formulaBytes, offset, externSheets, externalReferences, sheetNames, out reference)) {
                        failure = BiffFormulaReadFailure.Reference("FormulaDefinedName3dArea", "Defined-name 3D area could not be resolved.", token, tokenOffset);
                        return false;
                    }

                    return true;
                default:
                    failure = BiffFormulaReadFailure.UnsupportedToken(token, tokenOffset);
                    return false;
            }
        }

        internal static bool TryReadFormula(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formulaText,
            out BiffFormulaReadFailure? failure) {
            formulaText = null;
            failure = null;
            if (formulaBytes.Length == 0) {
                failure = BiffFormulaReadFailure.InvalidPayload("Defined-name formula token stream was empty.");
                return false;
            }

            byte[] payload = PrefixFormulaLength(formulaBytes);
            return BiffFormulaTextReader.TryRead(
                payload,
                parsedFormulaOffset: 0,
                formulaRow: 0,
                formulaColumn: 0,
                externSheets,
                externalReferences,
                sheetNames,
                definedNames,
                out formulaText,
                out failure);
        }

        internal static bool TryReadPrintTitles(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryReadPrintTitles(formulaBytes, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out reference);
        }

        internal static bool TryReadPrintTitles(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryReadPrintTitles(formulaBytes, externSheets, externalReferences, sheetNames, out reference, out _);
        }

        internal static bool TryReadPrintTitles(
            byte[] formulaBytes,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference,
            out BiffFormulaReadFailure? failure) {
            reference = null;
            failure = null;
            if (formulaBytes.Length == 0) {
                failure = BiffFormulaReadFailure.InvalidPayload("Print-title defined-name formula token stream was empty.");
                return false;
            }

            var parts = new List<string>();
            int unionCount = 0;
            int offset = 0;
            while (offset < formulaBytes.Length) {
                int tokenOffset = offset;
                byte token = formulaBytes[offset++];
                if (Is3dAreaToken(token)) {
                    if (!BiffFormulaReferenceFormatter.TryRead3dAreaInfo(formulaBytes, offset, externSheets, externalReferences, sheetNames, out BiffFormulaAreaReference area)) {
                        failure = offset + 10 > formulaBytes.Length
                            ? BiffFormulaReadFailure.InvalidPayload("Print-title 3D area formula ended early.", token, tokenOffset)
                            : BiffFormulaReadFailure.Reference("FormulaPrintTitle3dArea", "Print-title 3D area could not be resolved.", token, tokenOffset);
                        return false;
                    }

                    string? titleReference = FormatPrintTitleArea(area);
                    if (titleReference == null) {
                        failure = BiffFormulaReadFailure.Reference("FormulaPrintTitleAreaShape", "Print-title area is not a whole-row, whole-column, or rectangular title range.", token, tokenOffset);
                        return false;
                    }

                    parts.Add(titleReference);
                    offset += 10;
                    continue;
                }

                if (token == 0x10) {
                    if (parts.Count < 2) {
                        failure = BiffFormulaReadFailure.Stack("Print-title union token appeared before two title ranges.", token, tokenOffset);
                        return false;
                    }

                    unionCount++;
                    continue;
                }

                failure = BiffFormulaReadFailure.UnsupportedToken(token, tokenOffset);
                return false;
            }

            if (parts.Count == 0 || (parts.Count > 1 && unionCount != parts.Count - 1)) {
                failure = BiffFormulaReadFailure.Stack("Print-title formula tokens are individually recognized but the union expression could not be reduced.");
                return false;
            }

            reference = string.Join(",", parts);
            return true;
        }

        private static string? FormatPrintTitleArea(BiffFormulaAreaReference area) {
            int firstColumn = area.FirstColumnBits & 0x3fff;
            int lastColumn = area.LastColumnBits & 0x3fff;
            bool wholeRows = firstColumn == 0 && lastColumn == 255;
            bool wholeColumns = area.FirstRow == 0 && area.LastRow == ushort.MaxValue;

            if (wholeRows) {
                return area.SheetQualifier
                    + "!$"
                    + (area.FirstRow + 1).ToString(CultureInfo.InvariantCulture)
                    + ":$"
                    + (area.LastRow + 1).ToString(CultureInfo.InvariantCulture);
            }

            if (wholeColumns) {
                return area.SheetQualifier
                    + "!$"
                    + A1.ColumnIndexToLetters(firstColumn + 1)
                    + ":$"
                    + A1.ColumnIndexToLetters(lastColumn + 1);
            }

            string start = BiffFormulaReferenceFormatter.FormatCellReference(area.FirstRow, area.FirstColumnBits);
            string end = BiffFormulaReferenceFormatter.FormatCellReference(area.LastRow, area.LastColumnBits);
            return area.SheetQualifier + "!" + (start == end ? start : start + ":" + end);
        }

        private static bool Is3dAreaToken(byte token) {
            return token == 0x3b || token == 0x5b || token == 0x7b;
        }

        private static byte[] PrefixFormulaLength(byte[] formulaBytes) {
            byte[] payload = new byte[checked(formulaBytes.Length + 2)];
            payload[0] = (byte)(formulaBytes.Length & 0xff);
            payload[1] = (byte)((formulaBytes.Length >> 8) & 0xff);
            Buffer.BlockCopy(formulaBytes, 0, payload, 2, formulaBytes.Length);
            return payload;
        }
    }
}
