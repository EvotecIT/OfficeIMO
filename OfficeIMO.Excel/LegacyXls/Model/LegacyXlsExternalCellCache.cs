namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserved external cell cache section from an XCT record and its CRN records.
    /// </summary>
    public sealed class LegacyXlsExternalCellCache {
        private readonly List<LegacyXlsExternalCachedCell> _cells = new();

        internal LegacyXlsExternalCellCache(int declaredCrnCount, int? sheetIndex, string? sheetName, bool linkValid) {
            DeclaredCrnCount = declaredCrnCount;
            SheetIndex = sheetIndex;
            SheetName = sheetName;
            LinkValid = linkValid;
        }

        /// <summary>
        /// Gets the absolute CRN record count declared by the XCT record.
        /// </summary>
        public int DeclaredCrnCount { get; }

        /// <summary>
        /// Gets the zero-based external sheet index declared by the XCT record, when present.
        /// </summary>
        public int? SheetIndex { get; }

        /// <summary>
        /// Gets the external sheet name resolved from the preceding SupBook record, when available.
        /// </summary>
        public string? SheetName { get; }

        /// <summary>
        /// Gets whether the XCT record marked the preceding supporting link as valid.
        /// </summary>
        public bool LinkValid { get; }

        /// <summary>
        /// Gets cached cells preserved from CRN records in this section.
        /// </summary>
        public IReadOnlyList<LegacyXlsExternalCachedCell> Cells => _cells;

        /// <summary>
        /// Gets the first zero-based row occupied by the cache, when cached cells exist.
        /// </summary>
        public int? FirstRow => TryGetBounds(out int firstRow, out _, out _, out _) ? firstRow : null;

        /// <summary>
        /// Gets the last zero-based row occupied by the cache, when cached cells exist.
        /// </summary>
        public int? LastRow => TryGetBounds(out _, out int lastRow, out _, out _) ? lastRow : null;

        /// <summary>
        /// Gets the first zero-based column occupied by the cache, when cached cells exist.
        /// </summary>
        public int? FirstColumn => TryGetBounds(out _, out _, out int firstColumn, out _) ? firstColumn : null;

        /// <summary>
        /// Gets the last zero-based column occupied by the cache, when cached cells exist.
        /// </summary>
        public int? LastColumn => TryGetBounds(out _, out _, out _, out int lastColumn) ? lastColumn : null;

        /// <summary>
        /// Gets the row span covered by the cache bounding range, when cached cells exist.
        /// </summary>
        public int? RowSpan => TryGetBounds(out int firstRow, out int lastRow, out _, out _) ? lastRow - firstRow + 1 : null;

        /// <summary>
        /// Gets the column span covered by the cache bounding range, when cached cells exist.
        /// </summary>
        public int? ColumnSpan => TryGetBounds(out _, out _, out int firstColumn, out int lastColumn) ? lastColumn - firstColumn + 1 : null;

        /// <summary>
        /// Gets the occupied cache range in zero-based row/column notation, when cached cells exist.
        /// </summary>
        public string? CellRange => TryGetBounds(out int firstRow, out int lastRow, out int firstColumn, out int lastColumn)
            ? $"R{firstRow}C{firstColumn}:R{lastRow}C{lastColumn}"
            : null;

        internal List<LegacyXlsExternalCachedCell> MutableCells => _cells;

        private bool TryGetBounds(out int firstRow, out int lastRow, out int firstColumn, out int lastColumn) {
            if (_cells.Count == 0) {
                firstRow = 0;
                lastRow = 0;
                firstColumn = 0;
                lastColumn = 0;
                return false;
            }

            firstRow = int.MaxValue;
            lastRow = int.MinValue;
            firstColumn = int.MaxValue;
            lastColumn = int.MinValue;
            foreach (LegacyXlsExternalCachedCell cell in _cells) {
                if (cell.Row < firstRow) {
                    firstRow = cell.Row;
                }

                if (cell.Row > lastRow) {
                    lastRow = cell.Row;
                }

                if (cell.Column < firstColumn) {
                    firstColumn = cell.Column;
                }

                if (cell.Column > lastColumn) {
                    lastColumn = cell.Column;
                }
            }

            return true;
        }
    }
}
