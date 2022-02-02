using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTables() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(3, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 12, "Number of paragraphs during creation in table is wrong. Current: " + wordTable.Paragraphs.Count);

                Assert.True(document.Tables.Count == 1, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx"))) {
                Assert.True(document.Tables.Count == 1, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");

                var wordTable = document.Tables[0];
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 12, "Number of paragraphs during load in table is wrong. Current: " + wordTable.Paragraphs.Count);

                WordTable wordTable2 = document.AddTable(5, 5);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";
                wordTable2.Rows[4].Cells[0].Paragraphs[0].Text = "Test 5";

                Assert.True(wordTable2.Paragraphs.Count == 25, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);

                wordTable2.Rows[4].Remove();
                Assert.True(wordTable2.RowsCount == 4);
                wordTable2.AddRow(2, 0);
                Assert.True(wordTable2.RowsCount == 6);
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "");
                wordTable2.Rows[4].Cells[0].Paragraphs[0].Text = "Test 5";
                Assert.True(wordTable2.Rows[3].CellsCount == 5);
                Assert.True(wordTable2.Rows[4].CellsCount == 5);

                wordTable2.Rows[4].Cells[4].Remove();
                Assert.True(wordTable2.Rows[4].CellsCount == 4);
                Assert.True(wordTable2.Rows[5].CellsCount == 5);
                Assert.True(wordTable2.Paragraphs.Count == 29, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);


                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx"))) {
                Assert.True(document.Tables.Count == 2, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong.");

                var wordTable1 = document.Tables[0];
                Assert.True(wordTable1.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable1.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable1.Paragraphs.Count == 12, "Number of paragraphs during creation in table is wrong. Current: " + wordTable1.Paragraphs.Count);

                var wordTable2 = document.Tables[1];
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "Test 5", "Text in table matches. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Paragraphs.Count == 29, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);



                document.Save();
            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithAllTableStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllTableStyles.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                var listOfTablesStyles = (WordTableStyle[])Enum.GetValues(typeof(WordTableStyle));
                foreach (var tableStyle in listOfTablesStyles) {
                    var paragraph = document.AddParagraph(tableStyle.ToString());
                    paragraph.ParagraphAlignment = JustificationValues.Center;

                    WordTable wordTable = document.AddTable(4, 4, tableStyle);
                    wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                    wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                    wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                    wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                    Assert.True(wordTable.Style == tableStyle, "Table style matches");

                    Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                    Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable.Paragraphs.Count);
                }

                Assert.True(document.Tables.Count == 105, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 105, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 105, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[0].Tables.Count == 105, "Number of paragraphs on 1st section is wrong.");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllTableStyles.docx"))) {
                Assert.True(document.Tables.Count == 105, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 105, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 105, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[0].Tables.Count == 105, "Number of paragraphs on 1st section is wrong.");

                // lets read all tables and check their styles
                var listOfTablesStyles = (WordTableStyle[])Enum.GetValues(typeof(WordTableStyle));
                int count = 0;
                foreach (var tableStyle in listOfTablesStyles) {
                    WordTable loadedWordTable = document.Tables[count];

                    Assert.True(loadedWordTable.Rows.Count == 4, "Row count matches");
                    Assert.True(loadedWordTable.Rows[0].Cells.Count == 4, "Cells count matches");
                    Assert.True(loadedWordTable.Style == tableStyle, "Table style matches during load");
                    Assert.True(loadedWordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + loadedWordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                    Assert.True(loadedWordTable.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + loadedWordTable.Paragraphs.Count);

                    count++;
                }

                var wordTable = document.Tables[0];
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during load in table is wrong. Current: " + wordTable.Paragraphs.Count);

                WordTable wordTable2 = document.AddTable(5, 5);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";
                wordTable2.Rows[4].Cells[0].Paragraphs[0].Text = "Test 5";

                Assert.True(document.Tables.Count == 106, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 105, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 105, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[0].Tables.Count == 106, "Number of paragraphs on 1st section is wrong.");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllTableStyles.docx"))) {
                Assert.True(document.Tables.Count == 106, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 105, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 105, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[0].Tables.Count == 106, "Number of paragraphs on 1st section is wrong.");

                var wordTable1 = document.Tables[0];
                Assert.True(wordTable1.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable1.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable1.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable1.Paragraphs.Count);

                var wordTable2 = document.Tables[105];
                Assert.True(wordTable2.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table matches. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Paragraphs.Count == 25, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);

                document.Save();
            }
        }
    }
}
