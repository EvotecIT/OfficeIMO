using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
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
                            { "Q", "Revenue", "Churn" },
                            { "Q1", "1.1M", "2.1%" },
                            { "Q2", "1.3M", "1.8%" }
                        }).HeaderRow(0))
                    .End()
                    .Save(false);
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
        public void Test_FluentTableBuilder_Create() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentCreate.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t.Create(2, 2).Table!.Rows[1].Cells[1].AddParagraph("B"))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                var table = document.Tables[0];
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("B", table.Rows[1].Cells[1].Paragraphs[1].Text);
            }
        }

        [Fact]
        public void Test_FluentTableBuilder_AdvancedOperations() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTableBuilderAdvanced.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Create(2, 3)
                        .ForEachCell((r, c, cell) => cell.AddParagraph($"R{r}C{c}", true))
                        .Cell(1, 3, cell => cell.AddParagraph("Last", true))
                        .InsertRow(3, "A", "B", "C")
                        .InsertColumn(4, "X", "Y", "Z")
                        .RowStyle(1, r => r.Cells.ForEach(c => c.ShadingFillColorHex = "ffcccc"))
                        .ColumnStyle(2, c => c.ShadingFillColorHex = "ccffcc")
                        .Merge(1, 1, 2, 2)
                        .DeleteRow(3)
                        .DeleteColumn(4))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal(3, table.Rows[0].CellsCount);
                Assert.Equal("Last", table.Rows[0].Cells[2].Paragraphs[0].Text);
                Assert.Equal("R2C3", table.Rows[1].Cells[2].Paragraphs[0].Text);
                Assert.Equal("ffcccc", table.Rows[0].Cells[0].ShadingFillColorHex);
                Assert.Equal("ccffcc", table.Rows[0].Cells[1].ShadingFillColorHex);
                Assert.True(table.Rows[0].Cells[0].HasHorizontalMerge);
                Assert.True(table.Rows[0].Cells[0].HasVerticalMerge);
            }
        }

        [Fact]
        public void Test_FluentTableBuilder_PreferredWidthPoints() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTableBuilderPoints.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t.PreferredWidth(points: 100).Columns(2).Row("A", "B"))
                    .End()
                    .Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                var table = document.Tables[0];
                Assert.Equal(2000, table.Width);
                Assert.Equal(TableWidthUnitValues.Dxa, table.WidthType);
            }
        }

        [Fact]
        public void TableBuilderCellEnforces1BasedIndexing() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentTableBuilderInvalidCell.docx");
            using (var document = WordDocument.Create(filePath)) {
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Table(t => t.Cell(0, 1, _ => { })));
                Assert.Throws<ArgumentOutOfRangeException>(() => document.AsFluent().Table(t => t.Cell(1, 0, _ => { })));
            }
        }
    }
}

