using System.Globalization;
using System.Text;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffFormulaReferenceFormatter {
        internal const string MissingProjectedSheetReference = "#REF";

        internal static bool TryRead3dReference(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryRead3dReference(formulaBytes, offset, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out reference);
        }

        internal static bool TryRead3dReference(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            reference = null;
            if (offset + 6 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, externalReferences, sheetNames, externSheetIndex, out string? sheetQualifier)) {
                return false;
            }

            ushort row = BiffRecordReader.ReadUInt16(formulaBytes, offset + 2);
            ushort columnBits = BiffRecordReader.ReadUInt16(formulaBytes, offset + 4);
            reference = sheetQualifier + "!" + FormatCellReference(row, columnBits);
            return true;
        }

        internal static bool TryRead3dArea(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryRead3dArea(formulaBytes, offset, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out reference);
        }

        internal static bool TryRead3dArea(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            reference = null;
            if (!TryRead3dAreaInfo(formulaBytes, offset, externSheets, externalReferences, sheetNames, out BiffFormulaAreaReference area)) {
                return false;
            }

            string start = FormatCellReference(area.FirstRow, area.FirstColumnBits);
            string end = FormatCellReference(area.LastRow, area.LastColumnBits);
            reference = area.SheetQualifier + "!" + (start == end ? start : start + ":" + end);
            return true;
        }

        internal static bool TryRead3dInvalidReference(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            return TryRead3dInvalidReference(formulaBytes, offset, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out reference);
        }

        internal static bool TryRead3dInvalidReference(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            reference = null;
            if (offset + 2 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, externalReferences, sheetNames, externSheetIndex, out string? sheetQualifier)) {
                return false;
            }

            reference = sheetQualifier + "!#REF!";
            return true;
        }

        internal static bool TryReadExternalName(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            out string? reference) {
            reference = null;
            if (offset + 6 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            uint oneBasedNameIndex = BiffRecordReader.ReadUInt32(formulaBytes, offset + 2);
            if (externSheetIndex >= externSheets.Count || oneBasedNameIndex == 0 || oneBasedNameIndex > int.MaxValue) {
                return false;
            }

            ushort supBookIndex = externSheets[externSheetIndex].SupBookIndex;
            if (supBookIndex >= externalReferences.Count) {
                return false;
            }

            LegacyXlsExternalReference externalReference = externalReferences[supBookIndex];
            int nameIndex = checked((int)oneBasedNameIndex) - 1;
            if (nameIndex >= externalReference.ExternalNames.Count) {
                return false;
            }

            string name = externalReference.ExternalNames[nameIndex].Name;
            if (string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            reference = FormatExternalNameReference(externalReference, name);
            return true;
        }

        internal static bool TryRead3dAreaInfo(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out BiffFormulaAreaReference area) {
            return TryRead3dAreaInfo(formulaBytes, offset, externSheets, Array.Empty<LegacyXlsExternalReference>(), sheetNames, out area);
        }

        internal static bool TryRead3dAreaInfo(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            out BiffFormulaAreaReference area) {
            area = default;
            if (offset + 10 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, externalReferences, sheetNames, externSheetIndex, out string? sheetQualifier)) {
                return false;
            }

            area = new BiffFormulaAreaReference(
                sheetQualifier!,
                BiffRecordReader.ReadUInt16(formulaBytes, offset + 2),
                BiffRecordReader.ReadUInt16(formulaBytes, offset + 4),
                BiffRecordReader.ReadUInt16(formulaBytes, offset + 6),
                BiffRecordReader.ReadUInt16(formulaBytes, offset + 8));
            return true;
        }

        internal static string FormatCellReference(ushort zeroBasedRow, ushort columnBits) {
            int zeroBasedColumn = columnBits & 0x3fff;
            bool columnRelative = (columnBits & 0x4000) != 0;
            bool rowRelative = (columnBits & 0x8000) != 0;
            return (columnRelative ? string.Empty : "$")
                + A1.ColumnIndexToLetters(zeroBasedColumn + 1)
                + (rowRelative ? string.Empty : "$")
                + (zeroBasedRow + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static bool TryResolveSheetQualifier(
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            ushort externSheetIndex,
            out string? sheetQualifier) {
            sheetQualifier = null;
            if (externSheetIndex >= externSheets.Count) {
                return false;
            }

            BiffExternSheetReference externSheet = externSheets[externSheetIndex];
            short firstSheetIndex = externSheet.FirstSheetIndex;
            short lastSheetIndex = externSheet.LastSheetIndex;
            if (TryResolveExternalWorkbookSheetQualifier(externSheet, externalReferences, firstSheetIndex, lastSheetIndex, out sheetQualifier)) {
                return true;
            }

            if (firstSheetIndex < 0
                || lastSheetIndex < firstSheetIndex
                || lastSheetIndex >= sheetNames.Count) {
                return false;
            }

            if (sheetNames[firstSheetIndex] == MissingProjectedSheetReference || sheetNames[lastSheetIndex] == MissingProjectedSheetReference) {
                sheetQualifier = MissingProjectedSheetReference;
                return true;
            }

            string sheetReference = firstSheetIndex == lastSheetIndex
                ? sheetNames[firstSheetIndex]
                : sheetNames[firstSheetIndex] + ":" + sheetNames[lastSheetIndex];
            sheetQualifier = QuoteSheetReference(sheetReference);
            return true;
        }

        private static string FormatExternalNameReference(LegacyXlsExternalReference externalReference, string name) {
            if (externalReference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook
                && !string.IsNullOrWhiteSpace(externalReference.Target)) {
                return QuoteSheetReference(NormalizeExternalWorkbookTarget(externalReference.Target)) + "!" + name;
            }

            return name;
        }

        private static bool TryResolveExternalWorkbookSheetQualifier(
            BiffExternSheetReference externSheet,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            short firstSheetIndex,
            short lastSheetIndex,
            out string? sheetQualifier) {
            sheetQualifier = null;
            if (externSheet.SupBookIndex >= externalReferences.Count) {
                return false;
            }

            LegacyXlsExternalReference externalReference = externalReferences[externSheet.SupBookIndex];
            bool workbookLikeReference = externalReference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook
                || externalReference.Kind == LegacyXlsExternalReferenceKind.Unused;
            if (!workbookLikeReference
                || firstSheetIndex < 0
                || lastSheetIndex < firstSheetIndex
                || lastSheetIndex >= externalReference.SheetNames.Count) {
                return false;
            }

            string sheetReference = firstSheetIndex == lastSheetIndex
                ? externalReference.SheetNames[firstSheetIndex]
                : externalReference.SheetNames[firstSheetIndex] + ":" + externalReference.SheetNames[lastSheetIndex];
            sheetQualifier = QuoteSheetReference("[" + NormalizeExternalWorkbookTarget(externalReference.Target) + "]" + sheetReference);
            return true;
        }

        private static string NormalizeExternalWorkbookTarget(string? target) {
            if (string.IsNullOrWhiteSpace(target)) {
                return "ExternalWorkbook";
            }

            string sanitized = RemoveControlCharacters(target!);
            int separatorIndex = Math.Max(sanitized.LastIndexOf('\\'), sanitized.LastIndexOf('/'));
            if (separatorIndex >= 0 && separatorIndex + 1 < sanitized.Length) {
                sanitized = sanitized.Substring(separatorIndex + 1);
            }

            return string.IsNullOrWhiteSpace(sanitized) ? "ExternalWorkbook" : sanitized;
        }

        private static string RemoveControlCharacters(string value) {
            StringBuilder? sanitized = null;
            for (int i = 0; i < value.Length; i++) {
                if (!char.IsControl(value[i])) {
                    sanitized?.Append(value[i]);
                    continue;
                }

                sanitized ??= new StringBuilder(value.Length).Append(value, 0, i);
            }

            return sanitized == null ? value : sanitized.ToString();
        }

        private static string QuoteSheetReference(string sheetReference) {
            return "'" + sheetReference.Replace("'", "''") + "'";
        }
    }

    internal readonly struct BiffFormulaAreaReference {
        internal BiffFormulaAreaReference(string sheetQualifier, ushort firstRow, ushort lastRow, ushort firstColumnBits, ushort lastColumnBits) {
            SheetQualifier = sheetQualifier;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumnBits = firstColumnBits;
            LastColumnBits = lastColumnBits;
        }

        internal string SheetQualifier { get; }

        internal ushort FirstRow { get; }

        internal ushort LastRow { get; }

        internal ushort FirstColumnBits { get; }

        internal ushort LastColumnBits { get; }
    }
}
