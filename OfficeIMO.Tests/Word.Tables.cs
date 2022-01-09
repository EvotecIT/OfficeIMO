using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Color = System.Drawing.Color;
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

                var paragraph = document.InsertParagraph("Basic paragraph - Page 4");
                paragraph.ParagraphAlignment = JustificationValues.Center;

                WordTable wordTable = document.AddTable(3, 4);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";

                Assert.True(wordTable.Rows[2].Cells[0].Paragraphs[0].Text == "Test 3", "Text in table matches. Actual text: " + wordTable.Rows[2].Cells[0].Paragraphs[0].Text);

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

                WordTable wordTable2 = document.AddTable(5, 5);
                wordTable2.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";
                wordTable2.Rows[1].Cells[0].Paragraphs[0].Text = "Test 2";
                wordTable2.Rows[2].Cells[0].Paragraphs[0].Text = "Test 3";
                wordTable2.Rows[3].Cells[0].Paragraphs[0].Text = "Test 4";
                wordTable2.Rows[4].Cells[0].Paragraphs[0].Text = "Test 5";

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
                
                var wordTable2 = document.Tables[1];
                Assert.True(wordTable2.Rows[4].Cells[0].Paragraphs[0].Text == "Test 5", "Text in table matches. Actual text: " + wordTable2.Rows[4].Cells[0].Paragraphs[0].Text);

                document.Save();
            }
        }
    }
}
