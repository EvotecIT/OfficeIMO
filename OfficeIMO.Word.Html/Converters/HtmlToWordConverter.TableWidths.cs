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
                : Math.Max(1, (int)(checked((uint?)(section.PageSettings.Width?.ToTwips())) ?? checked((uint)(WordPageSizes.A4.Width.ToTwips())))
                    - (int)checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) - (int)checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips())));
            int desired = available;
            if (table.Width?.LengthValue is { } length && length > global::OfficeIMO.OfficeLength.Zero) desired = checked((int)length.ToTwips());
            else if (table.Width?.PercentValue is double percent && percent > 0)
                desired = checked((int)Math.Round(available * percent / 100d));
            if (cell != null) desired = Math.Min(desired, available);
            int width = Math.Max(1, Math.Min(2400, desired / columns));
            // Authored column and cell widths are applied afterward.
            if (width < 2400) table.ColumnWidth = Enumerable.Repeat<WordTableWidth?>(WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus((long)width * 635)), columns).ToList();
        }
    }
}
