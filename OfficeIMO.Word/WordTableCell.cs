using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableCell {
        internal TableCell _tableCell;
        internal TableCellProperties _tableCellProperties;

        public List<WordParagraph> Paragraphs {
            get {
                var list = new List<WordParagraph>();
                foreach (Paragraph paragraph in _tableCell.ChildElements.OfType<Paragraph>().ToList()) {
                    WordParagraph wordParagraph = new WordParagraph(_document, paragraph, null);
                    list.Add(wordParagraph);
                }

                return list;
            }
        }
        private readonly WordTable _wordTable;
        private readonly WordTableRow _wordTableRow;
        private readonly WordDocument _document;

        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow) {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" });

            // Specify the width property of the table cell.
            tableCell.Append(tableCellProperties);

            // Specify the table cell content.
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));

            WordParagraph paragraph = new WordParagraph();
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));
            //Paragraphs.Add(paragraph);

            tableCell.Append(paragraph._paragraph);

            wordTableRow._tableRow.Append(tableCell);

            _tableCellProperties = tableCellProperties;
            _tableCell = tableCell;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;
        }

        public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow, TableCell tableCell) {
            _tableCell = tableCell;
            _tableCellProperties = tableCell.TableCellProperties;
            _wordTable = wordTable;
            _wordTableRow = wordTableRow;
            _document = document;

            //foreach (Paragraph paragraph in tableCell.ChildElements.OfType<Paragraph>().ToList()) {
            //    WordParagraph wordParagraph = new WordParagraph(document, paragraph, null);
            //    this.Paragraphs.Add(wordParagraph);
            //}
        }

        public void Remove() {
            _tableCell.Remove();
            //_wordTableRow.Cells.Remove(this);

        }
    }
}
