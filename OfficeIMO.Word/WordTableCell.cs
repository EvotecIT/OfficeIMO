using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class WordTableCell {
        internal TableCell _tableCell;
        internal TableCellProperties _tableCellProperties;

        public List<WordParagraph> Paragraphs = new List<WordParagraph>();

        public WordTableCell() {
            TableCell tableCell = new TableCell();
            TableCellProperties tableCellProperties = new TableCellProperties(new TableCellWidth() {Type = TableWidthUnitValues.Dxa, Width = "2400"});

            // Specify the width property of the table cell.
            tableCell.Append(tableCellProperties);

            // Specify the table cell content.
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));
            
            WordParagraph paragraph = new WordParagraph();
            //tableCell.Append(new Paragraph(new Run(new Text("Hello, World!"))));
            Paragraphs.Add(paragraph);
            
            tableCell.Append(paragraph._paragraph);
       
            _tableCellProperties = tableCellProperties;
            _tableCell = tableCell;
        }

        public WordTableCell(TableCell tableCell, WordDocument document) {
            _tableCell = tableCell;
            _tableCellProperties = tableCell.TableCellProperties;

            foreach (Paragraph paragraph in tableCell.ChildElements.OfType<Paragraph>().ToList()) {
                WordParagraph wordParagraph = new WordParagraph(document, paragraph, null);
                this.Paragraphs.Add(wordParagraph);
            }
        }
    }
}
