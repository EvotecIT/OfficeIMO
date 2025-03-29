using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithNestedTables() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithNestedTables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Lets add table ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                WordTable wordTable1 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 1.2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 1.3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 1.4";
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1.5";


                WordTable wordTable2 = document.AddTable(5, 5, WordTableStyle.GridTable1LightAccent1);
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2.2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 2.3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 2.4";
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 2.5";

                var table1 = wordTable.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table2 = wordTable.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);


                var table3 = wordTable2.Rows[0].Cells[0].AddTable(3, 2, WordTableStyle.GridTable2Accent2);

                var table4 = wordTable2.Rows[0].Cells[1].AddTable(3, 2, WordTableStyle.GridTable2Accent5, true);

                Assert.True(document.Tables[0].HasNestedTables);
                Assert.False(document.Tables[0].IsNestedTable);

                Assert.True(table1.IsNestedTable);
                Assert.True(table2.IsNestedTable);

                wordTable.NestedTables[0].Rows[0].Cells[0].Paragraphs[0].Text = "Nested table 1 / 1st row / 1st cell";

                wordTable.NestedTables[1].Rows[1].Cells[1].Paragraphs[0].Text = "Nested table 2 / 2nd row / 2nd cell";

                Assert.True(table1.Rows[0].Cells[0].Paragraphs[0].Text == "Nested table 1 / 1st row / 1st cell");
                Assert.True(table2.Rows[1].Cells[1].Paragraphs[0].Text == "Nested table 2 / 2nd row / 2nd cell");

                Assert.True(document.Tables[0].HasNestedTables);
                Assert.False(document.Tables[0].IsNestedTable);

                Assert.True(document.Tables[0].NestedTables[0].IsNestedTable);
                Assert.True(document.Tables[0].NestedTables[1].IsNestedTable);
                Assert.True(document.Tables[0].NestedTables.Count == 2);
                Assert.True(document.Tables[1].NestedTables.Count == 0);
                Assert.True(document.Tables[2].NestedTables.Count == 2);
                Assert.True(document.Tables.Count == 3);
                Assert.True(document.TablesIncludingNestedTables.Count == 7);
                Assert.True(document.Sections[0].TablesIncludingNestedTables.Count == 7);

                Assert.True(table1.ParentTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2");
                Assert.True(table3.ParentTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2.2");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithNestedTables.docx"))) {
                Assert.True(document.Tables[0].HasNestedTables);
                Assert.False(document.Tables[0].IsNestedTable);

                Assert.True(document.Tables[0].NestedTables[0].IsNestedTable);
                Assert.True(document.Tables[0].NestedTables[1].IsNestedTable);
                Assert.True(document.Tables[0].NestedTables.Count == 2);
                Assert.True(document.Tables[1].NestedTables.Count == 0);
                Assert.True(document.Tables[2].NestedTables.Count == 2);
                Assert.True(document.Tables.Count == 3);
                Assert.True(document.TablesIncludingNestedTables.Count == 7);

                foreach (var table in document.TablesIncludingNestedTables) {
                    if (table.IsNestedTable) {
                        Assert.True(table.ParentTable.RowsCount > 0);
                    } else {
                        Assert.True(table.ParentTable == null);
                    }
                }

                foreach (var table in document.Sections[0].TablesIncludingNestedTables) {
                    if (table.IsNestedTable) {
                        Assert.True(table.ParentTable.RowsCount > 0);
                    } else {
                        Assert.True(table.ParentTable == null);
                    }
                }
                Assert.True(document.Sections[0].TablesIncludingNestedTables.Count == 7);

                document.Save();
            }
        }

        [Fact]
        public void Test_ReadingWordDocumentWithNestedTables() {
            string filePath = Path.Combine(_directoryDocuments, "NestedTables.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Single(document.Tables);
                
                var table = document.Tables[0];
                Assert.True(table.HasNestedTables);
                Assert.Equal(2, table.NestedTables.Count);
                Assert.Equal(9, table.Cells.Count);
                Assert.False(table.Cells[1].HasNestedTables);
                Assert.True(table.Cells[8].HasNestedTables);

                var cell = table.Cells[0];
                Assert.True(cell.HasNestedTables);
                Assert.Single(cell.NestedTables);
            }
        }
    }
}
