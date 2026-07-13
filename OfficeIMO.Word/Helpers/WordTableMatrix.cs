namespace OfficeIMO.Word {
    /// <summary>
    /// Maps native Word tables to a stable row/cell matrix for first-party adapters.
    /// </summary>
    internal static class WordTableMatrix {
        /// <summary>
        /// Maps an existing <see cref="WordTable"/> to a simple matrix of cells.
        /// This is useful for consumers that need to iterate over table cells
        /// without dealing with the underlying OpenXML structure.
        /// </summary>
        /// <param name="table">Table to map.</param>
        /// <returns>A matrix representing the table where each entry is a <see cref="WordTableCell"/>.</returns>
        internal static IEnumerable<IReadOnlyList<WordTableCell>> Map(WordTable table) {
            if (table == null) throw new ArgumentNullException(nameof(table));

            List<IReadOnlyList<WordTableCell>> result = new();
            foreach (WordTableRow row in table.Rows) {
                List<WordTableCell> cells = new();
                foreach (WordTableCell cell in row.Cells) {
                    cells.Add(cell);
                }
                result.Add(cells);
            }
            return result;
        }
    }
}
