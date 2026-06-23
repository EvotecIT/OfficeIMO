using System.Globalization;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffFormulaReferenceFormatter {
        internal static bool TryRead3dReference(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out string? reference) {
            reference = null;
            if (offset + 6 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, sheetNames, externSheetIndex, out string? sheetQualifier)) {
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
            reference = null;
            if (!TryRead3dAreaInfo(formulaBytes, offset, externSheets, sheetNames, out BiffFormulaAreaReference area)) {
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
            reference = null;
            if (offset + 2 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, sheetNames, externSheetIndex, out string? sheetQualifier)) {
                return false;
            }

            reference = sheetQualifier + "!#REF!";
            return true;
        }

        internal static bool TryRead3dAreaInfo(
            byte[] formulaBytes,
            int offset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            out BiffFormulaAreaReference area) {
            area = default;
            if (offset + 10 > formulaBytes.Length) {
                return false;
            }

            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaBytes, offset);
            if (!TryResolveSheetQualifier(externSheets, sheetNames, externSheetIndex, out string? sheetQualifier)) {
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
            IReadOnlyList<string> sheetNames,
            ushort externSheetIndex,
            out string? sheetQualifier) {
            sheetQualifier = null;
            if (externSheetIndex >= externSheets.Count) {
                return false;
            }

            short firstSheetIndex = externSheets[externSheetIndex].FirstSheetIndex;
            short lastSheetIndex = externSheets[externSheetIndex].LastSheetIndex;
            if (firstSheetIndex < 0
                || lastSheetIndex < firstSheetIndex
                || lastSheetIndex >= sheetNames.Count) {
                return false;
            }

            string sheetReference = firstSheetIndex == lastSheetIndex
                ? sheetNames[firstSheetIndex]
                : sheetNames[firstSheetIndex] + ":" + sheetNames[lastSheetIndex];
            sheetQualifier = QuoteSheetReference(sheetReference);
            return true;
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
