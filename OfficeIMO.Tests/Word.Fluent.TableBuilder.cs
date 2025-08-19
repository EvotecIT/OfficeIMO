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
                        .Columns(3).PreferredWidth(Percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(WordHorizontalAlignmentValues.Center))
                    .Table(t => t
                        .From2D(new object[,] {
                            { "Q", "Revenue", "Churn" },
                            { "Q1", "1.1M", "2.1%" },
                            { "Q2", "1.3M", "1.8%" }
                        }).HeaderRow(0))
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
        public void Test_FluentTableBuilder_AddTable() {
            string filePath = Path.Combine(_directoryWithFiles, "FluentAddTable.docx");
            using (var document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t.AddTable(2, 2).Table!.Rows[1].Cells[1].AddParagraph("B"))
                    .End();
                document.Save(false);
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                var table = document.Tables[0];
                Assert.Equal(2, table.Rows.Count);
                Assert.Equal("B", table.Rows[1].Cells[1].Paragraphs[1].Text);
            }
        }
    }
}

