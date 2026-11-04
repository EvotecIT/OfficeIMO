using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests paragraph indentation helpers.
    /// </summary>
    public partial class Word {
        /// <summary>
        /// Validates indentation values set in points.
        /// </summary>
        [Fact]
        public void Test_ParagraphIndentationPoints() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithIndentationPoints.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Indented paragraph");
                paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(10D);
                paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(5D);
                paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(15D);
                paragraph.IndentationHanging = global::OfficeIMO.OfficeLength.Points(2D);

                Assert.Equal(10, paragraph.IndentationBefore?.ToPoints());
                Assert.Equal(5, paragraph.IndentationAfter?.ToPoints());
                Assert.Equal(15, paragraph.IndentationFirstLine?.ToPoints());
                Assert.Equal(2, paragraph.IndentationHanging?.ToPoints());

                Assert.Equal(200, checked((int?)(paragraph.IndentationBefore?.ToTwips())));
                Assert.Equal(100, checked((int?)(paragraph.IndentationAfter?.ToTwips())));
                Assert.Equal(300, checked((int?)(paragraph.IndentationFirstLine?.ToTwips())));
                Assert.Equal(40, checked((int?)(paragraph.IndentationHanging?.ToTwips())));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(10, p.IndentationBefore?.ToPoints());
                Assert.Equal(5, p.IndentationAfter?.ToPoints());
                Assert.Equal(15, p.IndentationFirstLine?.ToPoints());
                Assert.Equal(2, p.IndentationHanging?.ToPoints());
                Assert.Equal(200, checked((int?)(p.IndentationBefore?.ToTwips())));
                Assert.Equal(100, checked((int?)(p.IndentationAfter?.ToTwips())));
                Assert.Equal(300, checked((int?)(p.IndentationFirstLine?.ToTwips())));
                Assert.Equal(40, checked((int?)(p.IndentationHanging?.ToTwips())));
            }
        }
    }
}
