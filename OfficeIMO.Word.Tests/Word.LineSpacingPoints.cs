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
                p.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(10D);
                p.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12D);

                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBefore?.ToPoints());
                Assert.Equal(12, p.LineSpacingAfter?.ToPoints());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBefore?.ToPoints());
                Assert.Equal(12, p.LineSpacingAfter?.ToPoints());
                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(200, checked((int?)(p.LineSpacingBefore?.ToTwips())));
                Assert.Equal(240, checked((int?)(p.LineSpacingAfter?.ToTwips())));
            }
        }

        [Fact]
        public void Test_LineSpacingPoints_SetPointsGetTwips() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentLineSpacingPoints.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var p = document.AddParagraph("Points based spacing");
                p.LineSpacingPoints = 18;
                p.LineSpacingBefore = global::OfficeIMO.OfficeLength.Points(10D);
                p.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12D);

                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(WordLineSpacingRule.Exact, p.LineSpacingRule);
                Assert.Equal(200, checked((int?)(p.LineSpacingBefore?.ToTwips())));
                Assert.Equal(240, checked((int?)(p.LineSpacingAfter?.ToTwips())));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var p = document.Paragraphs[0];
                Assert.Equal(18, p.LineSpacingPoints);
                Assert.Equal(10, p.LineSpacingBefore?.ToPoints());
                Assert.Equal(12, p.LineSpacingAfter?.ToPoints());
                Assert.Equal(360, p.LineSpacing);
                Assert.Equal(WordLineSpacingRule.Exact, p.LineSpacingRule);
                Assert.Equal(200, checked((int?)(p.LineSpacingBefore?.ToTwips())));
                Assert.Equal(240, checked((int?)(p.LineSpacingAfter?.ToTwips())));
            }
        }

        [Fact]
        public void LineSpacingPointsKeepsAnExplicitMinimumRuleAndClearsTheValue() {
            using WordDocument document = WordDocument.Create();
            var paragraph = document.AddParagraph("Minimum line spacing");
            paragraph.LineSpacingRule = WordLineSpacingRule.AtLeast;
            paragraph.LineSpacingPoints = 18;
            Assert.Equal(WordLineSpacingRule.AtLeast, paragraph.LineSpacingRule);
            Assert.Equal(360, paragraph.LineSpacing);
            paragraph.LineSpacingPoints = null;
            Assert.Null(paragraph.LineSpacing);
            Assert.Null(paragraph.LineSpacingPoints);
            Assert.Equal(WordLineSpacingRule.AtLeast, paragraph.LineSpacingRule);
            using WordDocument reopened = WordDocument.Load(new MemoryStream(document.ToBytes()));
            Assert.Null(reopened.Paragraphs[0].LineSpacing);
            Assert.Equal(WordLineSpacingRule.AtLeast, reopened.Paragraphs[0].LineSpacingRule);
        }
    }
}
