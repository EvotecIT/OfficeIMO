using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithPageSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "CreateDocumentPageSettings.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.PageSettings.Orientation = PageOrientationValues.Landscape;
                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Landscape, "Page orientation should match");
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                document.AddParagraph("Section 0");
                document.Sections[0].PageSettings.PageSize = WordPageSize.A3;

                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Landscape);
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Landscape);

                document.AddSection();
                document.Sections[1].PageSettings.PageSize = WordPageSize.A4;
                document.AddParagraph("Section 1");

                Assert.True(document.Sections[1].PageSettings.Orientation == PageOrientationValues.Landscape);
                Assert.True(document.Sections[1].PageOrientation == PageOrientationValues.Landscape);

                document.AddSection();
                document.Sections[2].PageSettings.PageSize = WordPageSize.A5;
                document.AddParagraph("Section 2");

                Assert.True(document.Sections[2].PageSettings.Orientation == PageOrientationValues.Landscape);
                Assert.True(document.Sections[2].PageOrientation == PageOrientationValues.Landscape);

                document.AddSection();
                document.Sections[3].PageSettings.PageSize = WordPageSize.A6;
                document.Sections[3].PageOrientation = PageOrientationValues.Portrait;
                document.AddParagraph("Section 3");

                Assert.True(document.Sections[3].PageSettings.Orientation == PageOrientationValues.Portrait);
                Assert.True(document.Sections[3].PageOrientation == PageOrientationValues.Portrait);

                document.AddSection();
                document.Sections[4].PageSettings.PageSize = WordPageSize.Executive;
                document.Sections[4].PageSettings.Orientation = PageOrientationValues.Landscape;
                document.AddParagraph("Section 4");

                Assert.True(document.Sections[3].PageSettings.Orientation == PageOrientationValues.Portrait);
                Assert.True(document.Sections[3].PageOrientation == PageOrientationValues.Portrait);


                var section = document.AddSection();
                document.AddParagraph("Section 5");

                var section2 = document.AddSection();
                document.AddParagraph("Section 6");

                var section3 = document.AddSection();
                document.AddParagraph("Section 7");

                var section4 = document.AddSection();
                document.AddParagraph("Section 8");

                var section5 = document.AddSection();
                document.AddParagraph("Section 9");

                section2.PageSettings.PageSize = WordPageSize.Legal;
                section2.PageSettings.Width = 500;
                section.PageSettings.PageSize = WordPageSize.A3;
                section3.PageSettings.PageSize = WordPageSize.B5;
                section4.PageSettings.PageSize = WordPageSize.Letter;
                section5.PageSettings.PageSize = WordPageSize.Legal;

                Assert.True(document.Sections[1].PageSettings.PageSize == WordPageSize.A4);
                Assert.True(document.Sections[2].PageSettings.PageSize == WordPageSize.A5);

                Assert.True(section.Paragraphs[0].Text == "Section 5");
                Assert.True(section2.Paragraphs[0].Text == "Section 6");

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentPageSettings.docx"))) {
                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Landscape);

                document.Sections[0].PageSettings.Orientation = PageOrientationValues.Portrait;

                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Portrait);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentPageSettings.docx"))) {
                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Portrait);
                document.PageSettings.Orientation = PageOrientationValues.Landscape;
                Assert.True(document.PageSettings.Orientation == PageOrientationValues.Landscape);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreateDocumentPageSettings.docx"))) {
                Assert.True(document.Sections[0].PageSettings.Orientation == PageOrientationValues.Landscape);
            }
        }
    }
}
