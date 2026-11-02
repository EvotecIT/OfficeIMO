using System.IO;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;
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
                var line = paragraph.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Parse("#FF0000"), OfficeIMO.OfficeLength.Points(2));

                Assert.True(document.Paragraphs.Count == 1);
                Assert.NotNull(paragraph.Line);

                document.Save();
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
                var line = paragraph.AddLine(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(100), OfficeIMO.OfficeLength.Points(0), Color.Blue, OfficeIMO.OfficeLength.Points(2));

                Assert.Equal(Color.Blue.ToRgbHex(), line.ColorHex.Replace("#", ""));

                Assert.Equal(line.ColorHex.Replace("#", ""), line.ColorHex.Replace("#", "").ToUpperInvariant());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentWithLinesColor.docx"))) {
                Assert.True(document.Paragraphs[0].IsLine);
            }
        }
    }
}
