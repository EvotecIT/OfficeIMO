using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;

namespace OfficeIMO {
    public class WordTable {
        public List<WordTableRow> Rows = new List<WordTableRow>();
        internal Table _table;
        internal TableProperties _tableProperties;
        internal WordDocument _document;
        //internal string Text;
        internal WordSection _section;

        internal void GenerateTable(WordDocument document, WordSection section) {
            // Create an empty table.
            Table table = new Table();

            // Create a TableProperties object and specify its border information.
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 },
                    new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Dashed), Size = 24 }
                )
            );
            // Append the TableProperties object to the empty table.
            table.AppendChild<TableProperties>(tblProp);

            _document = document;
            _tableProperties = tblProp;
            _table = table;
            _section = section;
        }

        public WordTable(WordDocument document, WordSection section, Table table) {
            _table = table;
            _tableProperties = table.ChildElements.OfType<TableProperties>().FirstOrDefault();
            _document = document;
            _section = section;


            foreach (TableRow row in table.ChildElements.OfType<TableRow>().ToList()) {
                WordTableRow tableRow = new WordTableRow(row, document);
                this.Rows.Add(tableRow);
            }

            section.Tables.Add(this);
        }

        public WordTable(WordDocument document, WordSection section, int rows, int columns) {

            this.GenerateTable(document, section);
            
            //WordTable table = new WordTable(document, section);
            //this.Text = "TEst";
            for (int i = 0; i < rows; i++) {
                WordTableRow row = new WordTableRow();
                this.Add(row);
                for (int j = 0; j < columns; j++) {
                    WordTableCell cell = new WordTableCell();
                    row.Add(cell);
                }
            }

            //// Create a row.
            //TableRow tr = new TableRow();

            //// Create a cell.
            //TableCell tc1 = new TableCell();

            //// Specify the width property of the table cell.
            //tc1.Append(new TableCellProperties(new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = "2400" }));

            //// Specify the table cell content.
            //tc1.Append(new Paragraph(new Run(new Text("Hello, World!"))));

            //// Append the table cell to the table row.
            //tr.Append(tc1);

            //// Create a second table cell by copying the OuterXml value of the first table cell.
            //TableCell tc2 = new TableCell(tc1.OuterXml);

            //// Append the table cell to the table row.
            //tr.Append(tc2);

            //// Append the table row to the table.
            //table._table.Append(tr);

            // Append the table to the document.
            document._wordprocessingDocument.MainDocumentPart.Document.Body.Append(this._table);

            section.Tables.Add(this);
        }

        private void Add(WordTableRow row) {
            _table.Append(row._tableRow);
            this.Rows.Add(row);
        }
    }
}
