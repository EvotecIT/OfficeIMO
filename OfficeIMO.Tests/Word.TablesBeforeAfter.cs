using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithTablesAfterBefore() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAfterBefore.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                var paragraph = document.AddParagraph("Lets add table with some alignment ");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Bold = true;
                paragraph.Underline = UnderlineValues.DotDash;

                Assert.True(document.Paragraphs.Count == 1);

                WordTable wordTable = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 111";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                Assert.True(document.Tables.Count == 1);

                var paragraph1 = document.AddParagraph("Lets add another table showing text wrapping around, but notice table before and after it anyways, that we just added at the end of the document.");

                WordTable wordTable1 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable1.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable1.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable1.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable1.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";

                wordTable1.WidthType = TableWidthUnitValues.Pct;
                wordTable1.Width = 3000;

                wordTable1.AllowTextWrap = true;

                Assert.True(document.Paragraphs.Count == 2);
                Assert.True(document.Tables.Count == 2);

                var paragraph2 = document.AddParagraph("This paragraph should continue but next to to the table");

                document.AddParagraph();
                document.AddParagraph();

                Assert.True(document.Tables.Count == 2);
                Assert.True(document.Paragraphs.Count == 5);


                var paragraph3 = document.AddParagraph("Lets add another table showing AutoFit");

                WordTable wordTable2 = document.AddTable(4, 4, WordTableStyle.GridTable1LightAccent1);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";


                Assert.True(document.Tables.Count == 3);
                Assert.True(document.Paragraphs.Count == 6);


                paragraph1.AddParagraphBeforeSelf();
                paragraph1.AddParagraphAfterSelf();

                var table3 = paragraph1.AddTableAfter(4, 4, WordTableStyle.GridTable1LightAccent1);
                table3.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document after paragraph 1";

                Assert.True(table3.Rows[0].Cells[0].Paragraphs[0].Text == document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text);

                var table4 = paragraph1.AddTableBefore(4, 4, WordTableStyle.GridTable1LightAccent1);
                table4.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the middle of the document before paragraph 1";

                Assert.True(document.Tables.Count == 5);
                Assert.True(document.Paragraphs.Count == 8);

                Assert.True(table4.Rows[0].Cells[0].Paragraphs[0].Text == document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text);


                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAfterBefore.docx"))) {

                Assert.True(document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document before paragraph 1");
                Assert.True(document.Tables[2].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document after paragraph 1");
                Assert.True(document.Tables.Count == 5);

                var table0 = document.Paragraphs[0].AddTableBefore(3, 3);
                table0.Rows[0].Cells[0].Paragraphs[0].Text = "Inserted in the very beginning.";

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the very beginning.");
                Assert.True(document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text == "Test 111");
                Assert.True(document.Tables[2].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document before paragraph 1");
                Assert.True(document.Tables[3].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document after paragraph 1");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithTablesAfterBefore.docx"))) {

                Assert.True(document.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the very beginning.");
                Assert.True(document.Tables[1].Rows[0].Cells[0].Paragraphs[0].Text == "Test 111");
                Assert.True(document.Tables[2].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document before paragraph 1");
                Assert.True(document.Tables[3].Rows[0].Cells[0].Paragraphs[0].Text == "Inserted in the middle of the document after paragraph 1");

                document.Save();
            }
        }
    }
}
