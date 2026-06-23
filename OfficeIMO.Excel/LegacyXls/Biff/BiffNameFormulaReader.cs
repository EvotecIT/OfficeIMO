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
            reference = null;
            if (formulaBytes.Length == 0) {
                return false;
            }

            int offset = 0;
            byte token = formulaBytes[offset++];
            switch (token) {
                case 0x3a:
                case 0x5a:
                case 0x7a:
                    if (offset + 6 != formulaBytes.Length) return false;
                    return BiffFormulaReferenceFormatter.TryRead3dReference(formulaBytes, offset, externSheets, externalReferences, sheetNames, out reference);
                case 0x3b:
                case 0x5b:
                case 0x7b:
                    if (offset + 10 != formulaBytes.Length) return false;
                    return BiffFormulaReferenceFormatter.TryRead3dArea(formulaBytes, offset, externSheets, externalReferences, sheetNames, out reference);
                default:
                    return false;
            }
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
            reference = null;
            if (formulaBytes.Length == 0) {
                return false;
            }

            var parts = new List<string>();
            int unionCount = 0;
            int offset = 0;
            while (offset < formulaBytes.Length) {
                byte token = formulaBytes[offset++];
                if (Is3dAreaToken(token)) {
                    if (!BiffFormulaReferenceFormatter.TryRead3dAreaInfo(formulaBytes, offset, externSheets, externalReferences, sheetNames, out BiffFormulaAreaReference area)) {
                        return false;
                    }

                    string? titleReference = FormatPrintTitleArea(area);
                    if (titleReference == null) {
                        return false;
                    }

                    parts.Add(titleReference);
                    offset += 10;
                    continue;
                }

                if (token == 0x10) {
                    if (parts.Count < 2) {
                        return false;
                    }

                    unionCount++;
                    continue;
                }

                return false;
            }

            if (parts.Count == 0 || (parts.Count > 1 && unionCount != parts.Count - 1)) {
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
    }
}
