using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_DocumentCleanupFeature() {
            string filePath = Path.Combine(_directoryWithFiles, "SimpleWordDocumentReadyToCleanup.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddParagraph("This is a text ").AddText("more text").AddText(" even longer text").AddText(" and even longer right?");

                Assert.True(document.Paragraphs.Count == 4);

                // since WordParagraph above are actually "Runs" with the same formatting cleanup will merge them as a single WordParagraph (single Run)
                var changesCount = document.CleanupDocument();
                Assert.True(changesCount == 3);

                Assert.True(document.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs[0].Text == "This is a text more text even longer text and even longer right?");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentReadyToCleanup.docx"))) {
                Assert.True(document.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs[0].Text == "This is a text more text even longer text and even longer right?");

                document.AddParagraph("This is a text 1 ").AddText("more text 1").AddText(" even longer text 1").AddText(" and even longer right?");

                document.Paragraphs[3].Bold = true;
                document.Paragraphs[4].Bold = true;

                Assert.True(document.Paragraphs.Count == 5);

                document.CleanupDocument();

                Assert.True(document.Paragraphs.Count == 3);

                Assert.True(document.Paragraphs[0].Text == "This is a text more text even longer text and even longer right?");
                Assert.True(document.Paragraphs[1].Text == "This is a text 1 more text 1");
                Assert.True(document.Paragraphs[2].Text == " even longer text 1 and even longer right?");


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "SimpleWordDocumentReadyToCleanup.docx"))) {

                Assert.True(document.Paragraphs.Count == 3);

                document.Save(false);
            }
        }
    }
}
