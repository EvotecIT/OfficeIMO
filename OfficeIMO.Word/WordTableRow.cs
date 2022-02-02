using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableRow {
        internal readonly TableRow _tableRow;

        public List<WordTableCell> Cells {
            get {
                var list = new List<WordTableCell>();
                foreach (TableCell cell in _tableRow.ChildElements.OfType<TableCell>().ToList()) {
                    WordTableCell wordCell = new WordTableCell(_document, _wordTable, this, cell);
                    list.Add(wordCell);
                }

                return list;
            }
        }

        public int CellsCount => Cells.Count;
        private readonly WordTable _wordTable;
        private readonly WordDocument _document;

        public WordTableRow(WordDocument document, WordTable wordTable) {
            // Create a row.
            TableRow tableRow = new TableRow();
            _tableRow = tableRow;
            _document = document;
            _wordTable = wordTable;

        }
        public WordTableRow(WordTable wordTable, TableRow row, WordDocument document) {
            _document = document;
            _tableRow = row;
            _wordTable = wordTable;

            foreach (TableCell cell in row.ChildElements.OfType<TableCell>()) {
                WordTableCell wordCell = new WordTableCell(document, wordTable, this, cell);
                //this.Cells.Add(wordCell);
            }
        }
        //public void Add(WordTableCell cell) {
        //    _tableRow.Append(cell._tableCell);
        //    //this.Cells.Add(cell);
        //}
        public void Remove() {
            _tableRow.Remove();
            //_wordTable.Rows.Remove(this);
        }
    }
}
