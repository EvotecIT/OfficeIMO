using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_FluentTableBuilder() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTableBuilder.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Columns(3).PreferredWidth(percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(HorizontalAlignment.Center))
                    .Table(t => t
                        .From2D(new object[,] {
                            { "Q",  "Revenue", "Churn" },
                            { "Q1", "1.1M",    "2.1%" },
                            { "Q2", "1.3M",    "1.8%" }
                        })
                        .HeaderRow(1))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Tables.Count);

                var table1 = document.Tables[0];
                Assert.Equal(3, table1.Rows.Count);
                Assert.Equal(3, table1.Rows[0].CellsCount);
                Assert.Equal(5000, table1.Width);
                Assert.Equal(TableWidthUnitValues.Pct, table1.WidthType);
                Assert.Equal(WordTableStyle.TableGrid, table1.Style);
                Assert.Equal(TableRowAlignmentValues.Center, table1.Alignment);
                Assert.True(table1.ConditionalFormattingFirstRow);
                Assert.Equal("Name", table1.Rows[0].Cells[0].Paragraphs[0].Text);

                var table2 = document.Tables[1];
                Assert.Equal(3, table2.Rows.Count);
                Assert.True(table2.Rows[0].RepeatHeaderRowAtTheTopOfEachPage);
                Assert.Equal("Q1", table2.Rows[1].Cells[0].Paragraphs[0].Text);
            }
        }

        [Fact]
        public void Test_FluentTableBuilder_CreateTable() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentCreateTable.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Create(2, 2)
                        .Cell(2, 2).Text("B"))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                var table = document.Tables[0];
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("B", table.Rows[1].Cells[1].Paragraphs[0].Text);
            }
        }

        [Fact]
        public void Test_FluentTableBuilder_ExtendedOperations() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTableBuilderExtended.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Create(rows: 2, cols: 3)
                        .ForEachCell((r, c, cell) => cell.Text($"R{r}C{c}"))
                        .Cell(1, 3).Text("Done")
                        .InsertRow(2, "A", "B", "C")
                        .InsertColumn(2, "X", "Y", "Z")
                        .Row(1).EachCell(c => c.Shading("#ff0000"))
                        .Column(3).Shading("#00ff00")
                        .Merge(fromRow: 1, fromCol: 1, toRow: 2, toCol: 2)
                        .DeleteRow(2)
                        .DeleteColumn(2))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                var table = document.Tables[0];
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(3, table.Rows[0].CellsCount);
                Assert.Equal("R2C3", table.Rows[1].Cells[2].Paragraphs[0].Text);
                Assert.Equal(MergedCellValues.Restart, table.Rows[0].Cells[0].HorizontalMerge);
                Assert.Equal("ff0000", table.Rows[0].Cells[0].ShadingFillColorHex);
                Assert.Equal("00ff00", table.Rows[0].Cells[1].ShadingFillColorHex);
            }
        }
    }
}
