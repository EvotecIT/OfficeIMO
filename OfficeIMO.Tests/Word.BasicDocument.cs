using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_SimpleWordDocumentCreation() {
            var filePath = Path.Combine(_directoryWithFiles, "TestFileTemporary.docx");

            var path = File.Exists(filePath);
            File.Delete(filePath);
            Assert.False(path); // MUST BE FALSE

            WordDocument document = WordDocument.Create(filePath);

            document.Save();

            path = File.Exists(filePath);
            Assert.True(path);

            document.Dispose();
            File.Delete(filePath);
        }

        [Fact]
        public void Test_OpeningWordAndParagraphCountMatches() {
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryDocuments, "BasicDocument.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Paragraphs.Count == 12);

                // There is only one Table in this document.
                //Assert.True(document.Tables.Count() == 1);

                // This table has 12 Paragraphs.
                //Assert.True(t0.Paragraphs.Count() == 12);
            }
        }

        [Fact]
        public void Test_AllElements() {
            var docs = Directory.GetFiles(_directoryDocuments, "*.docx")
                .Where(doc => !Path.GetFileName(doc).StartsWith("~"))
                .ToArray();
            foreach (var doc in docs) {
                Console.WriteLine($"Processing document: {doc}");

                using (WordDocument document = WordDocument.Load(doc)) {
                    var allElements = document.Elements;
                    Assert.True(allElements.Count > 0, $"Document '{doc}' has no elements.");

                    var allElementsByType = document.ElementsByType;
                    Assert.True(allElementsByType.Count > 0, $"Document '{doc}' has no elements by type.");
                }
            }
        }
    }
}
