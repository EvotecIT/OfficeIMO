using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;
using System.Collections.Generic;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains table-related tests.
    /// </summary>
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

                Assert.True(wordTable.RepeatHeaderRowAtTheTopOfEachPage == false);

                foreach (var row in wordTable.Rows) {
                    Assert.True(row.RepeatHeaderRowAtTheTopOfEachPage == false);
                }

                wordTable.RepeatHeaderRowAtTheTopOfEachPage = true;

                Assert.True(wordTable.RepeatHeaderRowAtTheTopOfEachPage == true);

                foreach (var row in wordTable.Rows.Skip(1)) {
                    Assert.True(row.RepeatHeaderRowAtTheTopOfEachPage == false);
                }

                Assert.True(wordTable.Rows.Count == 3);

                wordTable.AddRow(2, 2);

                Assert.True(wordTable.Rows.Count == 5);

                Assert.True(wordTable.Rows[0].RepeatHeaderRowAtTheTopOfEachPage == true);

                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable.Paragraphs.Count);

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
                Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during load in table is wrong. Current: " + wordTable.Paragraphs.Count);

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

                wordTable2.Rows[2].Cells[2].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[2].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 1");
                wordTable2.Rows[2].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 2");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[0].Text == "Test 3");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[1].Text == "More text which means another paragraph 1");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Text == "More text which means another paragraph 2");

                wordTable2.Rows[2].Cells[2].Paragraphs[2].Text = "Change me";
                wordTable2.Rows[2].Cells[2].Paragraphs[2].SetColor(Color.Green);

                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Text == "Change me");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Color == Color.Green);


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
                Assert.True(wordTable1.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable1.Paragraphs.Count);

                var wordTable2 = document.Tables[1];
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "Test 5", "Text in table matches. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Paragraphs.Count == 31, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);



                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTablesAddRows() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                var paragraph = document.AddParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(1, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                WordTableRow row = wordTable.AddRow();
                row.Cells[0].Paragraphs[0].Text = "Test 2";

                List<WordTableRow> rows = wordTable.AddRow(2, 0);
                rows[0].Cells[0].Paragraphs[0].Text = "Test 3";
                rows[1].Cells[0].Paragraphs[0].Text = "Test 4";

                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2", "Text in table does not match. Actual text: " + wordTable.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table does not match. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table does not match. Actual text: " + wordTable.Rows[3].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable.Paragraphs.Count);

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
                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2", "Text in table does not match. Actual text: " + wordTable.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table does not match. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table does not match. Actual text: " + wordTable.Rows[3].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable.Paragraphs.Count == 16, "Number of paragraphs during load in table is wrong. Current: " + wordTable.Paragraphs.Count);

                WordTable wordTable2 = document.AddTable(1, 5);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                WordTableRow row = wordTable2.AddRow();
                row.Cells[0].Paragraphs[0].Text = "Test 2";

                List<WordTableRow> rows = wordTable2.AddRow(3, 0);
                rows[0].Cells[0].Paragraphs[0].Text = "Test 3";
                rows[1].Cells[0].Paragraphs[0].Text = "Test 4";
                rows[2].Cells[0].Paragraphs[0].Text = "Test 5";

                Assert.True(wordTable2.Paragraphs.Count == 25, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);
                Assert.True(wordTable2.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2", "Text in table does not match. Actual text: " + wordTable2.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table does not match. Actual text: " + wordTable2.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table does not match. Actual text: " + wordTable2.Rows[3].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "Test 5", "Text in table does not match. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);

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

                wordTable2.Rows[2].Cells[2].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[2].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 1");
                wordTable2.Rows[2].Cells[2].Paragraphs[0].AddText("More text which means another paragraph 2");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[0].Text == "Test 3");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[1].Text == "More text which means another paragraph 1");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Text == "More text which means another paragraph 2");

                wordTable2.Rows[2].Cells[2].Paragraphs[2].Text = "Change me";
                wordTable2.Rows[2].Cells[2].Paragraphs[2].SetColor(Color.Green);

                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Text == "Change me");
                Assert.True(wordTable2.Rows[2].Cells[2].Paragraphs[2].Color == Color.Green);


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
                Assert.True(wordTable1.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable1.Paragraphs.Count);

                var wordTable2 = document.Tables[1];

                Assert.True(wordTable2.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2", "Text in table does not match. Actual text: " + wordTable2.Rows[1].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table does not match. Actual text: " + wordTable2.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table does not match. Actual text: " + wordTable2.Rows[3].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "Test 5", "Text in table does not match. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Paragraphs.Count == 31, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);



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


                WordTable wordTable3 = document.AddTable(5, 5);

                WordTable wordTable4 = document.AddTable(5, 5);

                WordTable wordTable5 = document.AddTable(7, 6);

                wordTable4.Remove();

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithAllTableStyles.docx"))) {
                Assert.True(document.Tables.Count == 108, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");
                Assert.True(document.Paragraphs.Count == 105, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 105, "Number of paragraphs on 1st section is wrong.");
                Assert.True(document.Sections[0].Tables.Count == 108, "Number of paragraphs on 1st section is wrong.");

                var wordTable1 = document.Tables[0];
                Assert.True(wordTable1.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable1.Rows[2].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable1.Paragraphs.Count == 16, "Number of paragraphs during creation in table is wrong. Current: " + wordTable1.Paragraphs.Count);

                var wordTable2 = document.Tables[105];
                Assert.True(wordTable2.Rows[3].Cells[0].Paragraphs[0].Text == "Test 4", "Text in table matches. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);
                Assert.True(wordTable2.Paragraphs.Count == 25, "Number of paragraphs during creation in table is wrong. Current: " + wordTable2.Paragraphs.Count);

                var wordTable3 = document.Tables[107];
                Assert.True(wordTable3.RowsCount == 7);
                Assert.True(wordTable3.Rows[0].CellsCount == 6);

                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTablesWithMerging() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Some test";
                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "Some test 1";
                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable.Rows[0].Cells[3].Paragraphs[0].Text = "Some test 3";


                Assert.True(document.Tables.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs.Count == 1);

                wordTable.Rows[0].Cells[1].MergeHorizontally(2, true);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 3);
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs.Count == 1);

                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text == "Some test 1");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[1].Text == "Some test 2");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[2].Text == "Some test 3");

                // should be empty paragraphs
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs[0].Text == "");
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs[0].Text == "");

                Assert.True(wordTable.Rows[0].Cells[1].HorizontalMerge == MergedCellValues.Restart);
                Assert.True(wordTable.Rows[0].Cells[2].HorizontalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[0].Cells[3].HorizontalMerge == MergedCellValues.Continue);

                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.PlainTable1);

                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Some test";
                wordTable2.Rows[0].Cells[1].Paragraphs[0].Text = "Some test 1";
                wordTable2.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable2.Rows[0].Cells[3].Paragraphs[0].Text = "Some test 3";
                wordTable2.Rows[0].Cells[1].MergeHorizontally(2, false);

                Assert.True(document.Tables[1].Rows[0].Cells[1].Paragraphs.Count == 1);
                Assert.True(document.Tables[1].Rows[0].Cells[1].Paragraphs.Count == 1);
                Assert.True(document.Tables[1].Rows[0].Cells[1].Paragraphs.Count == 1);

                Assert.True(document.Tables[1].Rows[0].Cells[1].Paragraphs[0].Text == "Some test 1");

                // should be empty paragraphs
                Assert.True(document.Tables[1].Rows[0].Cells[2].Paragraphs[0].Text == "");
                Assert.True(document.Tables[1].Rows[0].Cells[3].Paragraphs[0].Text == "");




                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx"))) {
                var wordTable = document.Tables[0];

                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 3);
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text == "Some test 1");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[1].Text == "Some test 2");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[2].Text == "Some test 3");
                // should be empty paragraphs
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs[0].Text == "");
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs[0].Text == "");

                Assert.True(wordTable.Rows[0].Cells[1].HorizontalMerge == MergedCellValues.Restart);
                Assert.True(wordTable.Rows[0].Cells[2].HorizontalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[0].Cells[3].HorizontalMerge == MergedCellValues.Continue);

                document.Tables[0].Rows[0].Cells[1].SplitHorizontally(2);

                Assert.True(wordTable.Rows[0].Cells[1].HorizontalMerge == null);
                Assert.True(wordTable.Rows[0].Cells[2].HorizontalMerge == null);
                Assert.True(wordTable.Rows[0].Cells[3].HorizontalMerge == null);

                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 3);
                Assert.True(document.Tables[0].Rows[0].Cells[2].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[3].Paragraphs.Count == 1);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx"))) {

                document.AddParagraph("Another table");

                var wordTable = document.AddTable(7, 4, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 0";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "Some test 1";
                wordTable.Rows[2].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable.Rows[3].Cells[2].Paragraphs[0].Text = "Some test 3";
                wordTable.Rows[0].Cells[2].MergeVertically(2, true);

                Assert.True(wordTable.Rows[0].Cells[2].VerticalMerge == MergedCellValues.Restart);
                Assert.True(wordTable.Rows[1].Cells[2].VerticalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[2].Cells[2].VerticalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[3].Cells[2].VerticalMerge == null);
                Assert.True(wordTable.Rows[4].Cells[2].VerticalMerge == null);
                Assert.True(wordTable.Rows[0].Cells[2].Paragraphs[0].Text == "Some test 0");
                Assert.True(wordTable.Rows[0].Cells[2].Paragraphs[1].Text == "Some test 1");
                Assert.True(wordTable.Rows[0].Cells[2].Paragraphs[2].Text == "Some test 2");
                Assert.True(wordTable.Rows[1].Cells[2].Paragraphs[0].Text == "");
                Assert.True(wordTable.Rows[2].Cells[2].Paragraphs[0].Text == "");

                wordTable = document.AddTable(7, 4, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Some test 0";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "Some test 1";
                wordTable.Rows[2].Cells[2].Paragraphs[0].Text = "Some test 2";
                wordTable.Rows[3].Cells[2].Paragraphs[0].Text = "Some test 3";
                wordTable.Rows[0].Cells[2].MergeVertically(2, false);

                Assert.True(wordTable.Rows[0].Cells[2].VerticalMerge == MergedCellValues.Restart);
                Assert.True(wordTable.Rows[1].Cells[2].VerticalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[2].Cells[2].VerticalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[3].Cells[2].VerticalMerge == null);
                Assert.True(wordTable.Rows[4].Cells[2].VerticalMerge == null);
                Assert.True(wordTable.Rows[0].Cells[2].Paragraphs[0].Text == "Some test 0");
                Assert.True(wordTable.Rows[0].Cells[2].Paragraphs.Count == 1);
                Assert.True(wordTable.Rows[1].Cells[2].Paragraphs[0].Text == "");
                Assert.True(wordTable.Rows[2].Cells[2].Paragraphs[0].Text == "");


                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTables.docx"))) {
                var wordTable = document.Tables[2];

                // split vertically previously merged cells
                wordTable.Rows[0].Cells[2].SplitVertically(2);

                Assert.Null(wordTable.Rows[0].Cells[2].VerticalMerge);
                Assert.Null(wordTable.Rows[1].Cells[2].VerticalMerge);
                Assert.Null(wordTable.Rows[2].Cells[2].VerticalMerge);

                Assert.Equal("Some test 0", wordTable.Rows[0].Cells[2].Paragraphs[0].Text);
                Assert.Equal("Some test 1", wordTable.Rows[0].Cells[2].Paragraphs[1].Text);
                Assert.Equal("Some test 2", wordTable.Rows[0].Cells[2].Paragraphs[2].Text);
                Assert.Equal("", wordTable.Rows[1].Cells[2].Paragraphs[0].Text);
                Assert.Equal("", wordTable.Rows[2].Cells[2].Paragraphs[0].Text);

                document.Save();
            }
        }

        [Fact]
        public void Test_MergeVerticallyFirstColumnMissingProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "VerticalMergeMissingProperties.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable wordTable = document.AddTable(3, 1, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Top";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Bottom";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Another";

                // simulate cell loaded without properties
                wordTable.Rows[0].Cells[0].RemoveTableCellProperties();

                wordTable.Rows[0].Cells[0].MergeVertically(1, true);

                Assert.Equal(MergedCellValues.Restart, wordTable.Rows[0].Cells[0].VerticalMerge);
                Assert.Equal(MergedCellValues.Continue, wordTable.Rows[1].Cells[0].VerticalMerge);
                Assert.Null(wordTable.Rows[2].Cells[0].VerticalMerge);
                Assert.Equal("Top", wordTable.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("Bottom", wordTable.Rows[0].Cells[0].Paragraphs[1].Text);
                Assert.Equal("", wordTable.Rows[1].Cells[0].Paragraphs[0].Text);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "VerticalMergeMissingProperties.docx"))) {
                var wordTable = document.Tables[0];

                Assert.True(wordTable.Rows[0].Cells[0].VerticalMerge == MergedCellValues.Restart);
                Assert.True(wordTable.Rows[1].Cells[0].VerticalMerge == MergedCellValues.Continue);
                Assert.True(wordTable.Rows[2].Cells[0].VerticalMerge == null);
                Assert.True(wordTable.Rows[0].Cells[0].Paragraphs[0].Text == "Top");
                Assert.True(wordTable.Rows[0].Cells[0].Paragraphs[1].Text == "Bottom");
                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs[0].Text == "");

                document.Save();
            }
        }

        [Fact]
        public void Test_RowMergeVerticallyAddsProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "RowVerticalMerge.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable wordTable = document.AddTable(3, 3, WordTableStyle.PlainTable1);

                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Top";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Bottom";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Another";
                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "A";
                wordTable.Rows[1].Cells[1].Paragraphs[0].Text = "B";
                wordTable.Rows[2].Cells[1].Paragraphs[0].Text = "C";

                // simulate cells loaded without properties
                wordTable.Rows[0].Cells[0].RemoveTableCellProperties();
                wordTable.Rows[1].Cells[0].RemoveTableCellProperties();

                wordTable.Rows[0].MergeVertically(0, 1, true);

                Assert.Equal(MergedCellValues.Restart, wordTable.Rows[0].Cells[0].VerticalMerge);
                Assert.Equal(MergedCellValues.Continue, wordTable.Rows[1].Cells[0].VerticalMerge);
                Assert.Null(wordTable.Rows[2].Cells[0].VerticalMerge);
                Assert.Equal("Top", wordTable.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("Bottom", wordTable.Rows[0].Cells[0].Paragraphs[1].Text);
                Assert.Equal("", wordTable.Rows[1].Cells[0].Paragraphs[0].Text);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "RowVerticalMerge.docx"))) {
                var wordTable = document.Tables[0];

                Assert.Equal(MergedCellValues.Restart, wordTable.Rows[0].Cells[0].VerticalMerge);
                Assert.Equal(MergedCellValues.Continue, wordTable.Rows[1].Cells[0].VerticalMerge);
                Assert.Null(wordTable.Rows[2].Cells[0].VerticalMerge);

                document.Save();
            }
        }

        [Fact]
        public void Test_SplitVerticallyKeepsIndices() {
            string filePath = Path.Combine(_directoryWithFiles, "SplitVerticallyIndices.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(4, 2, WordTableStyle.PlainTable1);

                table.Rows[0].Cells[0].MergeVertically(2, true);

                Assert.True(table.Rows[0].Cells[0].HasVerticalMerge);
                Assert.True(table.Rows[1].Cells[0].HasVerticalMerge);
                Assert.True(table.Rows[2].Cells[0].HasVerticalMerge);

                Assert.Equal(4, table.RowsCount);
                foreach (var row in table.Rows) {
                    Assert.Equal(2, row.CellsCount);
                }

                table.Rows[0].Cells[0].SplitVertically(2);

                Assert.Equal(4, table.RowsCount);
                foreach (var row in table.Rows) {
                    Assert.Equal(2, row.CellsCount);
                }

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SplitVerticallyIndices.docx"))) {
                var table = document.Tables[0];

                Assert.Equal(4, table.RowsCount);
                foreach (var row in table.Rows) {
                    Assert.Equal(2, row.CellsCount);
                }
                Assert.False(table.Rows[0].Cells[0].HasVerticalMerge);
                Assert.False(table.Rows[1].Cells[0].HasVerticalMerge);
                Assert.False(table.Rows[2].Cells[0].HasVerticalMerge);

                document.Save();
            }
        }

        [Fact]
        public void Test_SplitHorizontallyKeepsIndices() {
            string filePath = Path.Combine(_directoryWithFiles, "SplitHorizontallyIndices.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 4, WordTableStyle.PlainTable1);

                table.Rows[0].Cells[1].MergeHorizontally(2, true);

                Assert.Equal(2, table.RowsCount);
                foreach (var row in table.Rows) {
                    Assert.Equal(4, row.CellsCount);
                }

                Assert.True(table.Rows[0].Cells[1].HasHorizontalMerge);
                Assert.True(table.Rows[0].Cells[2].HasHorizontalMerge);
                Assert.True(table.Rows[0].Cells[3].HasHorizontalMerge);

                table.Rows[0].Cells[1].SplitHorizontally(2);

                foreach (var row in table.Rows) {
                    Assert.Equal(4, row.CellsCount);
                }

                Assert.False(table.Rows[0].Cells[1].HasHorizontalMerge);
                Assert.False(table.Rows[0].Cells[2].HasHorizontalMerge);
                Assert.False(table.Rows[0].Cells[3].HasHorizontalMerge);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SplitHorizontallyIndices.docx"))) {
                var table = document.Tables[0];

                Assert.Equal(2, table.RowsCount);
                foreach (var row in table.Rows) {
                    Assert.Equal(4, row.CellsCount);
                }
                Assert.False(table.Rows[0].Cells[1].HasHorizontalMerge);
                Assert.False(table.Rows[0].Cells[2].HasHorizontalMerge);
                Assert.False(table.Rows[0].Cells[3].HasHorizontalMerge);

                document.Save();
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithTablesWithSections() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndSections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Sections.Count == 1);

                WordTable wordTable = document.AddTable(3, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                document.AddSection();

                wordTable = document.AddTable(5, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 5";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 6";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 7";


                wordTable = document.AddTable(7, 8);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 8";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 9";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 10";

                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Sections[0].Tables.Count == 1);
                Assert.True(document.Sections[1].Tables.Count == 2);
                Assert.True(document.Tables.Count == 3);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndSections.docx"))) {
                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Sections[0].Tables.Count == 1);
                Assert.True(document.Sections[1].Tables.Count == 2);
                Assert.True(document.Tables.Count == 3);

                WordTable wordTable = document.AddTable(3, 8);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 11";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 12";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 13";

                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Sections[0].Tables.Count == 1);
                Assert.True(document.Sections[1].Tables.Count == 3);
                Assert.True(document.Tables.Count == 4);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndSections.docx"))) {
                Assert.True(document.Sections.Count == 2);
                Assert.True(document.Sections[0].Tables.Count == 1);
                Assert.True(document.Sections[1].Tables.Count == 3);
                Assert.True(document.Tables.Count == 4);

                document.Save();
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTablesAndOptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndOptions.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable.FirstRow.FirstCell.Paragraphs[0].AddComment("Adam Kłys", "AK", "Test comment for paragraph within a Table");

                Assert.True(wordTable.FirstRow.FirstCell.ShadingFillColor == null);

                wordTable.FirstRow.FirstCell.ShadingFillColor = Color.Blue;

                //Assert.True(wordTable.FirstRow.FirstCell.Paragraphs[0].Comments.Count == 1);
                Assert.True(wordTable.FirstRow.FirstCell.ShadingFillColor == Color.Blue);


                wordTable.Rows[1].FirstCell.ShadingFillColor = Color.Red;

                Assert.True(wordTable.Rows[1].FirstCell.ShadingFillColor == Color.Red);

                Assert.True(wordTable.LastRow.FirstCell.ShadingPattern == null);

                wordTable.LastRow.FirstCell.ShadingPattern = ShadingPatternValues.Percent20;

                Assert.True(wordTable.LastRow.FirstCell.ShadingPattern == ShadingPatternValues.Percent20);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndOptions.docx"))) {

                var wordTable = document.Tables[0];

                Assert.True(wordTable.Rows[1].FirstCell.ShadingFillColor == Color.Red);
                Assert.True(wordTable.LastRow.FirstCell.ShadingPattern == ShadingPatternValues.Percent20);

                wordTable.Rows[1].FirstCell.ShadingFillColorHex = "#0000FF";

                Assert.True(wordTable.Rows[1].FirstCell.ShadingFillColor == Color.Blue);
                Assert.True(wordTable.Rows[1].FirstCell.ShadingFillColorHex == "0000ff");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndOptions.docx"))) {

                Assert.True(document.Tables[0].Rows[1].FirstCell.ShadingFillColor == Color.Blue);
                Assert.True(document.Tables[0].Rows[1].FirstCell.ShadingFillColorHex == "0000ff");

                document.Save();
            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithTablesAndMoreOptions() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndMoreOptions.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                Assert.True(wordTable1.Alignment == null);

                wordTable1.Alignment = TableRowAlignmentValues.Center;

                Assert.True(wordTable1.Alignment == TableRowAlignmentValues.Center);

                wordTable1.WidthType = TableWidthUnitValues.Pct;
                wordTable1.Width = 3000;

                wordTable1.Title = "This is a title of the table";
                wordTable1.Description = "This is a table showing some features";

                Assert.True(wordTable1.Description == "This is a table showing some features");
                Assert.True(wordTable1.Title == "This is a title of the table");

                Assert.True(wordTable1.AllowTextWrap == false);
                Assert.True(wordTable1.Position.VerticalAnchor == null);

                wordTable1.AllowTextWrap = true;

                Assert.True(wordTable1.AllowTextWrap == true);
                Assert.True(wordTable1.Position.VerticalAnchor == VerticalAnchorValues.Text);

                Assert.True(wordTable1.AllowOverlap == false);


                Assert.True(wordTable1.Position.TableOverlap == null);

                wordTable1.AllowOverlap = true;

                Assert.True(wordTable1.AllowOverlap == true);
                Assert.True(wordTable1.Position.TableOverlap == TableOverlapValues.Overlap);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndMoreOptions.docx"))) {
                var wordTable1 = document.Tables[0];

                // add a cell to 3rd row
                WordTableCell cell = new WordTableCell(document, wordTable1, wordTable1.Rows[2]);
                cell.Paragraphs[0].Text = "This cell is outside a bit";
                cell.TextDirection = TextDirectionValues.TopToBottomLeftToRightRotated;

                Assert.True(cell.TextDirection == TextDirectionValues.TopToBottomLeftToRightRotated);
                Assert.True(cell.Paragraphs[0].Text == "This cell is outside a bit");

                Assert.True(wordTable1.Rows[1].Cells.Count == 4);
                Assert.True(wordTable1.Rows[2].Cells.Count == 5);
                Assert.True(wordTable1.Rows[1].CellsCount == 4);
                Assert.True(wordTable1.Rows[2].Cells[4].Paragraphs[0].Text == "This cell is outside a bit");
                Assert.True(wordTable1.Rows[2].Cells[4].TextDirection == TextDirectionValues.TopToBottomLeftToRightRotated);

                Assert.True(wordTable1.Alignment == TableRowAlignmentValues.Center);

                Assert.True(wordTable1.AllowTextWrap == true);
                Assert.True(wordTable1.Position.VerticalAnchor == VerticalAnchorValues.Text);

                Assert.True(wordTable1.AllowOverlap == true);
                Assert.True(wordTable1.Position.TableOverlap == TableOverlapValues.Overlap);

                Assert.True(wordTable1.Description == "This is a table showing some features");
                Assert.True(wordTable1.Title == "This is a title of the table");

                Assert.True(wordTable1.Position.RightFromText == null);
                Assert.True(wordTable1.Position.LeftFromText == null);
                Assert.True(wordTable1.Position.TablePositionXAlignment == null);
                Assert.True(wordTable1.Position.TablePositionY == null);
                Assert.True(wordTable1.Position.HorizontalAnchor == null);


                wordTable1.Position.LeftFromText = 100;

                wordTable1.Position.RightFromText = 180;


                wordTable1.Position.TopFromText = 50;

                wordTable1.Position.BottomFromText = 130;

                wordTable1.Position.TablePositionXAlignment = HorizontalAlignmentValues.Left;

                wordTable1.Position.HorizontalAnchor = HorizontalAnchorValues.Margin;

                wordTable1.Position.TablePositionY = 1;

                document.Save();
            }


            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAndMoreOptions.docx"))) {
                var wordTable1 = document.Tables[0];

                Assert.True(wordTable1.Alignment == TableRowAlignmentValues.Center);

                Assert.True(wordTable1.AllowTextWrap == true);
                Assert.True(wordTable1.Position.VerticalAnchor == VerticalAnchorValues.Text);

                Assert.True(wordTable1.AllowOverlap == true);
                Assert.True(wordTable1.Position.TableOverlap == TableOverlapValues.Overlap);

                Assert.True(wordTable1.Description == "This is a table showing some features");
                Assert.True(wordTable1.Title == "This is a title of the table");

                Assert.True(wordTable1.Position.RightFromText == 180);
                Assert.True(wordTable1.Position.LeftFromText == 100);
                Assert.True(wordTable1.Position.TopFromText == 50);
                Assert.True(wordTable1.Position.BottomFromText == 130);
                Assert.True(wordTable1.Position.TablePositionXAlignment == HorizontalAlignmentValues.Left);
                Assert.True(wordTable1.Position.TablePositionY == 1);
                Assert.True(wordTable1.Position.HorizontalAnchor == HorizontalAnchorValues.Margin);

                Assert.True(wordTable1.Rows[1].Cells.Count == 4);
                Assert.True(wordTable1.Rows[2].Cells.Count == 5);
                Assert.True(wordTable1.Rows[1].CellsCount == 4);
                Assert.True(wordTable1.Rows[2].Cells[4].Paragraphs[0].Text == "This cell is outside a bit");
                Assert.True(wordTable1.Rows[2].Cells[4].TextDirection == TextDirectionValues.TopToBottomLeftToRightRotated);

            }
        }



        [Fact]
        public void Test_CreatingWordDocumentWithTablesAndSizes() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTablesAndSizes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Table 1");
                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable.Width == 0);
                Assert.True(wordTable.WidthType == TableWidthUnitValues.Auto);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 2 - Sized for 2000 width / Centered");
                WordTable wordTable1 = document.AddTable(2, 6, WordTableStyle.PlainTable1);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 1 - ok longer text, no autosize right?";
                wordTable1.WidthType = TableWidthUnitValues.Pct;
                wordTable1.Width = 2000;
                wordTable1.Alignment = TableRowAlignmentValues.Center;

                Assert.True(wordTable1.Width == 2000);
                Assert.True(wordTable1.WidthType == TableWidthUnitValues.Pct);
                Assert.True(wordTable1.Alignment == TableRowAlignmentValues.Center);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 3 - By default the table is autosized for full width");
                WordTable wordTable2 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable2.Width == 0);
                Assert.True(wordTable2.WidthType == TableWidthUnitValues.Auto);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 4 - Magic number 5000 (full width)");
                WordTable wordTable3 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable3.WidthType = TableWidthUnitValues.Pct;
                wordTable3.Width = 5000;
                wordTable3.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable3.Width == 5000);
                Assert.True(wordTable3.WidthType == TableWidthUnitValues.Pct);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 5 - 50% by using 2500 width (pct)");
                WordTable wordTable4 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable4.WidthType = TableWidthUnitValues.Pct;
                wordTable4.Width = 2500;
                wordTable4.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable4.Width == 2500);
                Assert.True(wordTable4.WidthType == TableWidthUnitValues.Pct);


                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 6 - 50% by using 2500 width (pct), that we fix to full width");
                WordTable wordTable5 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                // this data is temporary just to prove things work
                wordTable5.WidthType = TableWidthUnitValues.Pct;
                wordTable5.Width = 2500;
                // lets fix it for full width
                wordTable5.DistributeColumnsEvenly();

                // Verify the NEW behavior: DistributeColumnsEvenly distributes within the current width context.
                // Since Width was set to 2500 (Pct), it should remain 2500.
                Assert.True(wordTable5.Width == 2500, $"Width after DistributeColumnsEvenly should be 2500, but was {wordTable5.Width}");
                Assert.True(wordTable5.WidthType == TableWidthUnitValues.Pct, $"WidthType after DistributeColumnsEvenly should be Pct, but was {wordTable5.WidthType}");

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 6 - 50%");
                WordTable wordTable6 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable6.SetWidthPercentage(50);

                Assert.True(wordTable6.Width == 2500);
                Assert.True(wordTable6.WidthType == TableWidthUnitValues.Pct);

                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("Table 6 - 75%");
                WordTable wordTable7 = document.AddTable(3, 4, WordTableStyle.PlainTable1);
                wordTable7.SetWidthPercentage(75);

                Assert.True(wordTable7.Width == 3750);
                Assert.True(wordTable7.WidthType == TableWidthUnitValues.Pct);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTablesAndSizes.docx"))) {
                Assert.True(document.Tables[0].Width == 0);
                Assert.True(document.Tables[0].WidthType == TableWidthUnitValues.Auto);

                Assert.True(document.Tables[1].Width == 2000);
                Assert.True(document.Tables[1].WidthType == TableWidthUnitValues.Pct);

                Assert.True(document.Tables[2].Width == 0);
                Assert.True(document.Tables[2].WidthType == TableWidthUnitValues.Auto);

                Assert.True(document.Tables[3].Width == 5000);
                Assert.True(document.Tables[3].WidthType == TableWidthUnitValues.Pct);

                Assert.True(document.Tables[4].Width == 2500);
                Assert.True(document.Tables[4].WidthType == TableWidthUnitValues.Pct);

                // Verify loaded state for the table where DistributeColumnsEvenly was called
                Assert.True(document.Tables[5].Width == 2500, $"[Load] Width after DistributeColumnsEvenly should be 2500, but was {document.Tables[5].Width}");
                Assert.True(document.Tables[5].WidthType == TableWidthUnitValues.Pct, $"[Load] WidthType after DistributeColumnsEvenly should be Pct, but was {document.Tables[5].WidthType}");

                Assert.True(document.Tables[6].Width == 2500);
                Assert.True(document.Tables[6].WidthType == TableWidthUnitValues.Pct);

                Assert.True(document.Tables[7].Width == 3750);
                Assert.True(document.Tables[7].WidthType == TableWidthUnitValues.Pct);

            }
        }


        [Fact]
        public void Test_CreatingWordDocumentWithTables1() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTables1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "This paragraph preexists on table creation";

                wordTable.Rows[1].Cells[0].AddParagraph();
                wordTable.Rows[1].Cells[0].Paragraphs[1].Text = "Paragraph added separately to the text";

                wordTable.Rows[1].Cells[0].AddParagraph("Another paragraph within table cell, added directly");

                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs.Count == 3);

                WordParagraph paragraph = new WordParagraph {
                    Text = "Paragraph added separately as WordParagraph",
                    Bold = true,
                    Italic = true
                };

                wordTable.Rows[1].Cells[0].AddParagraph(paragraph);

                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs.Count == 4);

                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTablesCellAlignment() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatingWordDocumentWithTables1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                WordTable wordTable = document.AddTable(4, 3);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "This is the normal vertical alignment";
                wordTable.Rows[0].Cells[0].VerticalAlignment = TableVerticalAlignmentValues.Top;

                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "But it can be centered";
                wordTable.Rows[0].Cells[1].VerticalAlignment = TableVerticalAlignmentValues.Center;

                wordTable.Rows[0].Cells[2].Paragraphs[0].Text = "Or at the bottom";
                wordTable.Rows[0].Cells[2].VerticalAlignment = TableVerticalAlignmentValues.Bottom;

                Assert.True(wordTable.Rows[0].Cells[0].VerticalAlignment == TableVerticalAlignmentValues.Top);
                Assert.True(wordTable.Rows[0].Cells[1].VerticalAlignment == TableVerticalAlignmentValues.Center);
                Assert.True(wordTable.Rows[0].Cells[2].VerticalAlignment == TableVerticalAlignmentValues.Bottom);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var wordTable1 = document.Tables[0];
                Assert.True(wordTable1.Rows[0].Cells[0].VerticalAlignment == TableVerticalAlignmentValues.Top);
                Assert.True(wordTable1.Rows[0].Cells[1].VerticalAlignment == TableVerticalAlignmentValues.Center);
                Assert.True(wordTable1.Rows[0].Cells[2].VerticalAlignment == TableVerticalAlignmentValues.Bottom);

                document.Save();
            }
        }


        [Fact]
        public void Test_SetTableStyleId() {

            string originalFile = Path.Combine(_directoryDocuments, "CreatingWordDocumentWithTables2.docx");
            string tempFile = Path.GetTempFileName();

            using (WordDocument document = WordDocument.Create(originalFile)) {

                document.AddParagraph("Table1");

                WordTable wordTable1 = document.AddTable(4, 3);
                wordTable1.SetStyleId(WordTableStyle.TableNormal.ToString());

                document.AddParagraph("Table2");

                WordTable wordTable2 = document.AddTable(4, 3);
                wordTable2.Style = WordTableStyle.TableNormal;

                Assert.True(wordTable1.Style == wordTable2.Style);
                document.Save(false);

            }

        }


        [Fact]
        public void Test_CreatingWordDocumentWithTablesWithReplace() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesReplace.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddParagraph("Table1 - Test 1");
                document.AddParagraph("Test1");


                WordTable wordTable = document.AddTable(3, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                Assert.True(document.Tables.Count == 1);

                var wordTable = document.Tables[0];

                Assert.True(wordTable.Rows[0].Cells[0].Paragraphs[0].Text == "Test 1");
                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 2");
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3");

                WordDocument.FindAndReplace(wordTable.Rows[0].Cells[0].Paragraphs, "Test 1", "Test 11");
                WordDocument.FindAndReplace(wordTable.Rows[1].Cells[0].Paragraphs, "Test 2", "Test 21");

                // lets make sure find and replace works only on table
                Assert.True(document.Paragraphs[0].Text == "Table1 - Test 1");
                Assert.True(document.Paragraphs[1].Text == "Test1");

                // lets check data for table
                Assert.True(wordTable.Rows[0].Cells[0].Paragraphs[0].Text == "Test 11");
                Assert.True(wordTable.Rows[1].Cells[0].Paragraphs[0].Text == "Test 21");
                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3");


                document.Save();
            }

        }



        [Fact]
        public void Test_CreatingWordDocumentWithTablesClearParagraphs() {
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

                wordTable.Rows[0].Cells[1].Paragraphs[0].Text = "Test Row 0 Cell 1";

                Assert.True(document.Tables.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 1);

                // add 2 more texts to the same cell as new paragraphs
                wordTable.Rows[0].Cells[0].Paragraphs[0].AddParagraph("New");
                wordTable.Rows[0].Cells[0].Paragraphs[1].AddParagraph("New more");

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs.Count == 3);

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text == "Test 1");
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[1].Text == "New");
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[2].Text == "New more");

                // replace existing paragraphs with single one
                wordTable.Rows[0].Cells[0].AddParagraph("New paragraph, delete rest", true);

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text == "New paragraph, delete rest");

                // lets try to add new paragraph to the same cell, using WordParagraph
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text == "Test Row 0 Cell 1");

                WordParagraph paragraph1 = new WordParagraph {
                    Text = "Paragraph added separately as WordParagraph",
                    Bold = true,
                    Italic = true
                };

                wordTable.Rows[0].Cells[1].AddParagraph(paragraph1);

                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs.Count == 2);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[0].Text == "Test Row 0 Cell 1");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[1].Text == "Paragraph added separately as WordParagraph");
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[1].Bold == true);
                Assert.True(document.Tables[0].Rows[0].Cells[1].Paragraphs[1].Italic == true);

                document.Save(false);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithTableInsertCopyRow() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesInsertRow.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Tables.Count == 0, "Tables count matches");
                Assert.True(document.Lists.Count == 0, "List count matches");

                WordTable wordTable = document.AddTable(3, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2.1";
                wordTable.Rows[1].Cells[1].Paragraphs[0].Text = "Test 2.2";
                wordTable.Rows[1].Cells[2].Paragraphs[0].Text = "Test 2.3";

                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                Assert.True(document.Tables.Count == 1);
                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs.Count == 1);

                var newRow = wordTable.Rows[1];

                wordTable.CopyRow(newRow);

                Assert.True(document.Tables[0].Rows.Count == 4);
                Assert.True(document.Tables[0].Rows[3].Cells[0].Paragraphs[0].Text == "Test 2.1");
                Assert.True(document.Tables[0].Rows[3].Cells[1].Paragraphs[0].Text == "Test 2.2");
                Assert.True(document.Tables[0].Rows[3].Cells[2].Paragraphs[0].Text == "Test 2.3");

                wordTable.AddRow(4);

                Assert.True(document.Tables[0].Rows.Count == 5);

                wordTable.AddRow(3, 4);
                Assert.True(document.Tables[0].Rows.Count == 8);

                document.Save(false);
            }
        }

        [Fact]
        public void Test_CloningTable() {
            string filePath = Path.Combine(_directoryWithFiles, "TableClone.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

                WordTable cloned = table.Clone();

                Assert.Equal(2, document.Tables.Count);
                Assert.Equal("A1", cloned.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("B2", cloned.Rows[1].Cells[1].Paragraphs[0].Text);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Tables.Count);
                Assert.Equal(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text,
                             document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal(document.Tables[0].Rows[1].Cells[1].Paragraphs[0].Text,
                             document.Tables[1].Rows[1].Cells[1].Paragraphs[0].Text);
            }
        }

        [Fact]
        public void Test_TableRowHeightWithAutoFit() {
            string filePath = Path.Combine(_directoryWithFiles, "TableRowHeightAutoFit.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Height = 500;
                table.Rows[1].Height = 300;
                table.AutoFitToWindow();
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTable table = document.Tables[0];
                Assert.Equal(500, table.Rows[0].Height);
                Assert.Equal(300, table.Rows[1].Height);
            }
        }

    }
}
