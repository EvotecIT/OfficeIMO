namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void InitializeDefaultTableColumnWidths(WordTable table, int columns,
            WordDocument document, WordSection section, WordTableCell? cell) {
            if (columns <= 0 || table.Rows.Count == 0) return;
            // Use the conversion scope: list-item tables can already be inside an
            // SDT wrapper, which hides their cell or section from the native table.
            int available = cell != null
                ? (_tableCellContentWidths.TryGetValue(cell._tableCell, out int? cached) ? cached :
                    WordTable.EstimateCellContentWidthInDxa(document, cell._tableCell)) ?? 1
                : Math.Max(1, (int)(section.PageSettings.Width ?? WordPageSizes.A4.WidthTwips)
                    - (int)section.Margins.Left - (int)section.Margins.Right);
            int desired = available;
            if (table.WidthType == WordTableWidthUnit.Dxa && table.Width > 0) desired = table.Width.Value;
            else if (table.WidthType == WordTableWidthUnit.Pct && table.Width > 0)
                desired = (int)Math.Round(available * (double)table.Width.Value / 5000);
            if (cell != null) desired = Math.Min(desired, available);
            int width = Math.Max(1, Math.Min(2400, desired / columns));
            // Authored column and cell widths are applied afterward.
            if (width < 2400) table.ColumnWidth = Enumerable.Repeat(width, columns).ToList();
        }
    }
}
