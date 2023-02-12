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
                // this will need to be changed in the future to support this
                // TODO: add a method that searches and replaces over multiple WordParagraphs or even native paragraphs
                var replacedCount1 = document.FindAndReplace("This is a text more text", "Shorter text");
                Assert.True(replacedCount1 == 0);

                document.CleanupDocument();

                // cleanup should merge paragraphs making it easier to find and replace text
                // this only works for same formatting though
                // may require improvement in the future to ignore formatting completely, but then it's a bit tricky which formatting to apply
                var replacedCount2 = document.FindAndReplace("This is a text more text", "Shorter text");
                Assert.True(replacedCount2 == 1);

                var replacedCount3 = document.FindAndReplace("even longer", "not longer");
                Assert.True(replacedCount3 == 4);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {

                Assert.True(document.Paragraphs[0].Text == "Production Section");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentSearchFunctionality.docx"))) {



                document.Save(false);
            }
        }
    }
}
