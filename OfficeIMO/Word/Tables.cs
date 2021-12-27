using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public class Tables {
        public static void InsertTableInDoc(string filepath) {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(filepath, true)) {
                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // Create a table.
                Table tbl = new Table();

                // Set the style and width for the table.
                TableProperties tableProp = new TableProperties();
                DocumentFormat.OpenXml.Wordprocessing.TableStyle tableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle() {Val = "PlainTable5"};

                // Make the table width 100% of the page width.
                TableWidth tableWidth = new TableWidth() {Width = "5000", Type = TableWidthUnitValues.Pct};

                // Apply
                tableProp.Append(tableStyle, tableWidth);

                // Add 3 columns to the table.
                TableGrid tg = new TableGrid(new GridColumn(), new GridColumn(), new GridColumn());

                tbl.AppendChild(tg);

                // Create 1 row to the table.
                TableRow tr1 = new TableRow();

                // Add a cell to each column in the row.
                TableCell tc1 = new TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("1"))));
                TableCell tc2 = new TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("2"))));
                TableCell tc3 = new TableCell(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text("3"))));
                tr1.Append(tc1, tc2, tc3);

                // Add row to the table.
                tbl.AppendChild(tr1);

                // Add the table to the document
                body.AppendChild(tbl);
            }
        }


        public static void CreateWordprocessingDocument(string fileName) {
            string[,] data = {
                {"Texas", "TX"},
                {"California", "CA"},
                {"New York", "NY"},
                {"Massachusetts", "MA"}
            };

            using (var wordDocument = WordprocessingDocument.Open(fileName, true)) {
                // We need to change the file type from template to document.
                wordDocument.ChangeDocumentType(WordprocessingDocumentType.Document);

                StyleDefinitionsPart styles = wordDocument.MainDocumentPart.StyleDefinitionsPart;
                if (styles == null) {
                    // styles = AddStylesPartToPackage(wordDocument);

                    //CreateAndAddCharacterStyle(styles, "GridTable5Dark-Accent4",);
                    //CreateAndAddCharacterStyle(styles, "OverdueAmountChar", "Overdue Amount Char", "Late Due, Late Amount");
                    OfficeIMO.Word.WordDocument.AddDefaultStyleDefinitions(wordDocument, styles);
                }

                var body = wordDocument.MainDocumentPart.Document;

                Table table = new Table();

                TableProperties props = new TableProperties();
                DocumentFormat.OpenXml.Wordprocessing.TableStyle tableStyle = new DocumentFormat.OpenXml.Wordprocessing.TableStyle { Val = "GridTable4-Accent4"};
                props.Append(tableStyle);
                table.AppendChild(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++) {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++) {
                        var tc = new TableCell();
                        tc.Append(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(data[i, j]))));
                        tc.Append(new TableCellProperties(new TableCellWidth {Type = TableWidthUnitValues.Auto}));
                        tr.Append(tc);
                    }

                    table.Append(tr);
                }

                body.Append(table);
                wordDocument.MainDocumentPart.Document.Save();
            }
        }
    }
}