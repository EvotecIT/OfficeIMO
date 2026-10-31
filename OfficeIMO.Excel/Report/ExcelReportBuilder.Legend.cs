namespace OfficeIMO.Excel {
    // Legend block for ExcelReportBuilder.
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Renders a compact legend block with an optional title, column headers, and rows.
        /// Optionally colors the first column using a value→color mapping.
        /// </summary>
        /// <param name="title">Optional section title.</param>
        /// <param name="headers">Legend column headers.</param>
        /// <param name="rows">Legend row values.</param>
        /// <param name="firstColumnFillByValue">Optional first-column colors, using the dictionary's key comparer.</param>
        /// <param name="headerFill">Optional header fill; null uses the report key fill.</param>
        public ExcelReportBuilder SectionLegend(
            string? title,
            IReadOnlyList<string> headers,
            IEnumerable<IReadOnlyList<string>> rows,
            IDictionary<string, OfficeColor>? firstColumnFillByValue = null,
            OfficeColor? headerFill = null) {
            if (!string.IsNullOrWhiteSpace(title)) Section(title!);
            if (headers == null || headers.Count == 0) return this;
            OfficeColor fill = headerFill ?? _theme.KeyFill ?? _theme.Theme.Palette.Background;
            if (fill.ThemeSlot == null) fill = fill.Resolve(_theme.Theme);
            int headerRow = _row;
            for (int c = 0; c < headers.Count; c++) {
                Sheet.Cell(headerRow, c + 1, headers[c]);
                Sheet.CellBold(headerRow, c + 1, true);
                Sheet.CellBackground(headerRow, c + 1, fill);
            }
            _row++;

            if (rows != null) {
                foreach (var rvals in rows) {
                    int cols = System.Math.Min(rvals.Count, headers.Count);
                    for (int c = 0; c < cols; c++)
                        Sheet.Cell(_row, c + 1, rvals[c] ?? string.Empty);
                    if (cols > 0 && firstColumnFillByValue != null) {
                        if (Sheet.TryGetCellText(_row, 1, out var v) && v != null && firstColumnFillByValue.TryGetValue(v, out var color))
                            Sheet.CellBackground(_row, 1, color.ThemeSlot != null ? color : color.Resolve(_theme.Theme));
                    }
                    _row++;
                }
            }
            return Spacer();
        }
    }
}
