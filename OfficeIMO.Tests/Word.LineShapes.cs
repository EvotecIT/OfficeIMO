using System.IO;
using OfficeIMO.Word;
using SixLabors.ImageSharp;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains line shape tests.
    /// </summary>
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

        [Fact]
        public void Test_CreatingWordDocumentWithColorObject() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentWithLinesColor.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Paragraph with line");
                var line = paragraph.AddLine(0, 0, 100, 0, Color.Blue, 2);

                Assert.Equal(Color.Blue.ToHexColor(), line.ColorHex.Replace("#", ""));

                Assert.Equal(line.ColorHex.Replace("#", ""), line.ColorHex.Replace("#", "").ToLowerInvariant());

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithLinesColor.docx"))) {
                Assert.True(document.Paragraphs[0].IsLine);
            }
        }
    }
}
