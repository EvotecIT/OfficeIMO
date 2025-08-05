using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>
    /// Provides helper methods for constructing Wordprocessing tables from
    /// generic table representations.
    /// </summary>
    public static class TableBuilder {
        /// <summary>
        /// Builds a <see cref="Table"/> instance from the supplied structure.
        /// Each cell is represented by an action that populates a <see cref="TableCell"/>.
        /// </summary>
        /// <param name="structure">Table definition where the first dimension represents rows
        /// and the second dimension represents cells within a row.</param>
        /// <returns>Constructed <see cref="Table"/>.</returns>
        public static Table Build(IEnumerable<IEnumerable<Action<TableCell>>> structure) {
            if (structure == null) throw new ArgumentNullException(nameof(structure));

            Table table = new Table();
            foreach (IEnumerable<Action<TableCell>> rowDef in structure) {
                TableRow row = new TableRow();
                foreach (Action<TableCell> cellBuilder in rowDef) {
                    TableCell cell = new TableCell();
                    cellBuilder?.Invoke(cell);
                    if (!cell.HasChildren) {
                        cell.Append(new Paragraph());
                    }
                    row.Append(cell);
                }
                table.Append(row);
            }
            return table;
        }

        /// <summary>
        /// Maps an existing <see cref="WordTable"/> to a simple matrix of cells.
        /// This is useful for consumers that need to iterate over table cells
        /// without dealing with the underlying OpenXML structure.
        /// </summary>
        /// <param name="table">Table to map.</param>
        /// <returns>A matrix representing the table where each entry is a <see cref="WordTableCell"/>.</returns>
        public static IEnumerable<IReadOnlyList<WordTableCell>> Map(WordTable table) {
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
