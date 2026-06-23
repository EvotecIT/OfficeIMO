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

        internal List<LegacyXlsExternalCachedCell> MutableCells => _cells;
    }
}
