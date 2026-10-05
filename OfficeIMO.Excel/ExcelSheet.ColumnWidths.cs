namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Applies positive custom widths together and saves the worksheet XML once.
        /// Existing column styles, visibility and outline metadata are preserved.
        /// </summary>
        /// <param name="columnWidths">Widths keyed by 1-based column index. Widths above 255 are clamped.</param>
        /// <exception cref="ArgumentNullException">The width map is null.</exception>
        /// <exception cref="ArgumentOutOfRangeException">An index is outside the worksheet or a width is non-positive or non-finite.</exception>
        public void SetColumnWidths(IReadOnlyDictionary<int, double> columnWidths) {
            if (columnWidths == null) throw new ArgumentNullException(nameof(columnWidths));
            var selected = columnWidths.OrderBy(pair => pair.Key).ToArray();
            foreach (var pair in selected) {
                if (pair.Key < 1 || pair.Key > A1.MaxColumns || pair.Value <= 0D
                    || double.IsNaN(pair.Value) || double.IsInfinity(pair.Value)) {
                    throw new ArgumentOutOfRangeException(nameof(columnWidths),
                        "Column indexes must be within the worksheet and widths must be positive and finite.");
                }
            }
            if (selected.Length == 0) return;

            _excelDocument.MaterializeDeferredDataSetImport();
            WriteLock(() => {
                SetColumnWidthsCore(selected.Select(pair => pair.Key).ToArray(),
                    selected.Select(pair => pair.Value).ToArray());
                WorksheetRoot.Save();
            });
        }
    }
}
