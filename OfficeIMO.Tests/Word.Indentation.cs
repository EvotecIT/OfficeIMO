using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ParagraphIndentationPoints() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithIndentationPoints.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Indented paragraph");
                paragraph.IndentationBeforePoints = 10;
                paragraph.IndentationAfterPoints = 5;
                paragraph.IndentationFirstLinePoints = 15;
                paragraph.IndentationHangingPoints = 2;

                Assert.Equal(10, paragraph.IndentationBeforePoints);
                Assert.Equal(5, paragraph.IndentationAfterPoints);
                Assert.Equal(15, paragraph.IndentationFirstLinePoints);
                Assert.Equal(2, paragraph.IndentationHangingPoints);

                Assert.Equal(200, paragraph.IndentationBefore);
                Assert.Equal(100, paragraph.IndentationAfter);
                Assert.Equal(300, paragraph.IndentationFirstLine);
                Assert.Equal(40, paragraph.IndentationHanging);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(10, p.IndentationBeforePoints);
                Assert.Equal(5, p.IndentationAfterPoints);
                Assert.Equal(15, p.IndentationFirstLinePoints);
                Assert.Equal(2, p.IndentationHangingPoints);
                Assert.Equal(200, p.IndentationBefore);
                Assert.Equal(100, p.IndentationAfter);
                Assert.Equal(300, p.IndentationFirstLine);
                Assert.Equal(40, p.IndentationHanging);
            }
        }
    }
}
