using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class TableReadsKeepXmlTests {
        [Fact]
        public void ReadingBodyTablesLeavesCellsWithoutPropertiesUnchanged() {
            using var input = new MemoryStream();
            using (var source = WordDocument.Create()) {
                source.AddParagraph("Before");
                // Cells as other producers often write them: no tcPr, and a nested table without tcPr either.
                var nested = new Table(new TableGrid(new GridColumn()),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("Nested"))))));
                source._document.Body!.Append(new Table(new TableGrid(new GridColumn(), new GridColumn()),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("A1")))), new TableCell(new Paragraph(new Run(new Text("B1"))))),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text("A2"))), nested), new TableCell(new Paragraph(new Run(new Text("B2")))))));
                source.Save(input);
            }
            input.Position = 0;
            using var document = WordDocument.Load(input);
            Body body = document._document.Body!;
            Assert.Empty(body.Descendants<TableCellProperties>());
            string before = body.OuterXml;

            WordTable table = Assert.Single(document.Tables);
            Assert.Equal(2, table.RowCount);
            Assert.Equal("A1", table.FirstRow.FirstCell.Paragraphs[0].Text);
            Assert.Equal("B2", table.LastRow.LastCell.Paragraphs[0].Text);
            foreach (WordTableRow row in table.Rows) {
                Assert.Equal(2, row.CellsCount);
                foreach (WordTableCell cell in row.Cells) {
                    Assert.Null(cell.Width);
                    Assert.Null(cell.ShadingFillColor);
                    Assert.Null(cell.VerticalMerge);
                    Assert.Null(cell.Borders.LeftStyle);
                }
            }
            Assert.Equal(new[] { "A1", "B1", "A2", "B2" }, table.Paragraphs.Select(paragraph => paragraph.Text));
            Assert.Equal("Nested", Assert.Single(table.Rows[1].Cells[0].NestedTables).Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal(2, document.TablesIncludingNestedTables.Count);
            Assert.Equal(1, document.Find("B2").MatchCount);
            Assert.Empty(document.Images);
            Assert.Empty(document.ParagraphsImages);

            Assert.Equal(before, body.OuterXml);

            // A write still provisions tcPr, for the written cell only.
            table.Rows[0].Cells[1].ShadingFillColor = OfficeColor.Parse("#FF0000");
            Assert.Equal("FF0000", Assert.Single(body.Descendants<TableCellProperties>()).Shading!.Fill!.Value);
        }
    }
}
