namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Legend block for SheetComposer.
    /// </summary>
    public sealed partial class SheetComposer {
        /// <summary>
        /// Renders a compact legend block with an optional title, column headers, and rows.
        /// Optionally colors the first column using a value→color mapping.
        /// </summary>
        public SheetComposer SectionLegend(
            string? title,
            IReadOnlyList<string> headers,
            IEnumerable<IReadOnlyList<string>> rows,
            IDictionary<string, string>? firstColumnFillByValue = null,
            string? headerFillHex = null) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);
            if (headers == null || headers.Count == 0) return this;
            int headerRow = _row;
            for (int c = 0; c < headers.Count; c++) {
                Sheet.Cell(headerRow, c + 1, headers[c]);
                Sheet.CellBold(headerRow, c + 1, true);
                Sheet.CellBackground(headerRow, c + 1, headerFillHex ?? _theme.KeyFillHex);
            }
            _row++;

            if (rows != null) {
                foreach (var rvals in rows) {
                    int cols = System.Math.Min(rvals.Count, headers.Count);
                    for (int c = 0; c < cols; c++)
                        Sheet.Cell(_row, c + 1, rvals[c] ?? string.Empty);
                    if (cols > 0 && firstColumnFillByValue != null) {
                        if (Sheet.TryGetCellText(_row, 1, out var v) && v != null && firstColumnFillByValue.TryGetValue(v, out var hex))
                            Sheet.CellBackground(_row, 1, hex);
                    }
                    _row++;
                }
            }
            return Spacer();
        }
    }
}
