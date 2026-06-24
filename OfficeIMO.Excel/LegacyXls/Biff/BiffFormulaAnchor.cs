namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffFormulaAnchor {
        internal static bool TryGetFirstRangeAnchor(IReadOnlyList<string> ranges, out int zeroBasedRow, out int zeroBasedColumn) {
            zeroBasedRow = 0;
            zeroBasedColumn = 0;
            if (ranges.Count == 0) {
                return false;
            }

            string firstRange = ranges[0];
            if (A1.TryParseRange(firstRange, out int startRow, out int startColumn, out _, out _)
                || A1.TryParseCellReferenceFast(firstRange, out startRow, out startColumn)) {
                zeroBasedRow = startRow - 1;
                zeroBasedColumn = startColumn - 1;
                return true;
            }

            return false;
        }
    }
}
