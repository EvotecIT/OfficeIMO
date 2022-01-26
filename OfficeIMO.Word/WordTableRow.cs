using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableRow {
        internal TableRow _tableRow;

        public List<WordTableCell> Cells = new List<WordTableCell>();

        public WordTableRow() {
            // Create a row.
            TableRow tableRow = new TableRow();
            _tableRow = tableRow;
        }
        public WordTableRow(TableRow row, WordDocument document) {
            _tableRow = row;

            foreach (TableCell cell in row.ChildElements.OfType<TableCell>().ToList()) {
                WordTableCell wordCell = new WordTableCell(cell, document);
                this.Cells.Add(wordCell);
            }
        }

        public void Add(WordTableCell cell) {
            _tableRow.Append(cell._tableCell);
            this.Cells.Add(cell);
        }
    }
}
