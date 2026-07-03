using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_LineSpacingPoints_SetTwipsGetPoints() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentLineSpacingTwips.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p = document.AddParagraph("Twips based spacing");
                p.LineSpacing = 360;
                p.LineSpacingBefore = 200;
                p.LineSpacingAfter = 240;

                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBeforePoints);
                Assert.Equal(12, p.LineSpacingAfterPoints);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBeforePoints);
                Assert.Equal(12, p.LineSpacingAfterPoints);
                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(200, p.LineSpacingBefore);
                Assert.Equal(240, p.LineSpacingAfter);
            }
        }

        [Fact]
        public void Test_LineSpacingPoints_SetPointsGetTwips() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentLineSpacingPoints.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p = document.AddParagraph("Points based spacing");
                p.LineSpacingPoints = 18;
                p.LineSpacingBeforePoints = 10;
                p.LineSpacingAfterPoints = 12;

                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(200, p.LineSpacingBefore);
                Assert.Equal(240, p.LineSpacingAfter);

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBeforePoints);
                Assert.Equal(12, p.LineSpacingAfterPoints);
                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(200, p.LineSpacingBefore);
                Assert.Equal(240, p.LineSpacingAfter);
            }
        }
    }
}
