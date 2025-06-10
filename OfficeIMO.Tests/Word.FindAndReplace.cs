using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DocumentFindAndReplace() {
            string filePath = Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Test Section");

                document.Paragraphs[0].AddComment("Przemysław", "PK", "This is my comment");

                document.AddParagraph("Test Section - another line");

                document.Paragraphs[1].AddComment("Przemysław", "PK", "More comments");

                document.AddParagraph("This is a text ").AddText("more text").AddText(" even longer text").AddText(" and Even longer right?");

                document.AddParagraph("this is a text ").AddText("more text 1").AddText(" even longer text 1").AddText(" and Even longer right?");

                // we now ensure that we add bold to complicate the search
                document.Paragraphs[9].Bold = true;
                document.Paragraphs[10].Bold = true;

                var listFound = document.Find("Test Section");
                Assert.True(listFound.Count == 2);

                var replacedCount = document.FindAndReplace("Test Section", "Production Section");
                Assert.True(replacedCount == 2);

                // should be 0 because it stretches over 2 paragraphs
                var replacedCount1 = document.FindAndReplace("This is a text more text", "Shorter text");
                Assert.True(replacedCount1 == 2);

                document.CleanupDocument();

                // cleanup should merge paragraphs making it easier to find and replace text
                // this only works for same formatting though
                // may require improvement in the future to ignore formatting completely, but then it's a bit tricky which formatting to apply
                var replacedCount2 = document.FindAndReplace("This is a text more text", "Shorter text");
                Assert.True(replacedCount2 == 0);

                var replacedCount3 = document.FindAndReplace("even longer", "not longer");
                Assert.True(replacedCount3 == 4);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {

                Assert.True(document.Paragraphs[0].Text == "Production Section");

                var table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].AddText("Test Section");
                table.Rows[0].Cells[1].Paragraphs[0].AddText("Test Section");
                table.Rows[0].Cells[2].Paragraphs[0].AddText("Test ").AddText("Sect").AddText("ion");

                document.AddHeadersAndFooters();

                var header = document.Header.Default;
                var tableInHeader = document.AddTable(3, 3);
                tableInHeader.Rows[0].Cells[0].Paragraphs[0].AddText("Prod Section");
                tableInHeader.Rows[0].Cells[1].Paragraphs[0].AddText("Prod Section");
                tableInHeader.Rows[0].Cells[2].Paragraphs[0].AddText("Prod ").AddText("Sect").AddText("ion");

                var footer = document.Footer.Default;
                var tableInFooter = document.AddTable(3, 3);
                tableInFooter.Rows[0].Cells[0].Paragraphs[0].AddText("Prod Section");
                tableInFooter.Rows[0].Cells[1].Paragraphs[0].AddText("Prod Section");
                tableInFooter.Rows[0].Cells[2].Paragraphs[0].AddText("Prod ").AddText("Sect").AddText("ion");

                document.CleanupDocument();

                var listFound1 = document.Find("Test Section");
                Assert.True(listFound1.Count == 3);

                var listFound2 = document.Find("Prod Section");
                Assert.True(listFound2.Count == 6);

                var replacedCount = document.FindAndReplace("Prod Section", "Production Section");
                Assert.True(replacedCount == 6);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {



                document.Save(false);
            }
        }
    }
}
