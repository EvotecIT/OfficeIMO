using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithLines() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithLines.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with line");
                var line = paragraph.AddLine(0, 0, 100, 0, "#FF0000", 2);

                Assert.True(document.Paragraphs.Count == 1);
                Assert.NotNull(paragraph.Line);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithLines.docx"))) {
                Assert.True(document.Paragraphs[0].IsLine);
            }
        }
    }
}
