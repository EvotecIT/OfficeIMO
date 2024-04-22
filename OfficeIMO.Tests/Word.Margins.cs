using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithPageMargins2() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                document.AddParagraph("Section 0");
                document.Sections[0].Margins.Type = WordMargin.Normal;

                document.AddSection();
                document.Sections[1].Margins.Type = WordMargin.Narrow;
                document.AddParagraph("Section 1");

                document.AddSection();
                document.Sections[2].Margins.Type = WordMargin.Mirrored;
                document.AddParagraph("Section 2");

                document.AddSection();
                document.Sections[3].Margins.Type = WordMargin.Moderate;
                document.AddParagraph("Section 3");

                document.AddSection();
                document.Sections[4].Margins.Type = WordMargin.Wide;
                document.AddParagraph("Section 4");

                var section = document.AddSection();
                document.AddParagraph("Section 5");

                var section2 = document.AddSection();
                document.AddParagraph("Section 6");

                section2.Margins.Type = WordMargin.Office2003Default;
                section2.Margins.Bottom = 15;
                section.Margins.Type = WordMargin.Office2003Default;

                Assert.True(document.Sections[0].Margins.Type == WordMargin.Normal);
                Assert.True(document.Sections[1].Margins.Type == WordMargin.Narrow);
                Assert.True(document.Sections[2].Margins.Type == WordMargin.Mirrored);
                Assert.True(document.Sections[3].Margins.Type == WordMargin.Moderate);
                Assert.True(document.Sections[4].Margins.Type == WordMargin.Wide);
                Assert.True(document.Sections[5].Margins.Type == WordMargin.Office2003Default);
                Assert.True(document.Sections[6].Margins.Type == WordMargin.Unknown);

                Assert.True(section.Margins.Type == WordMargin.Office2003Default);
                Assert.True(section2.Margins.Type == WordMargin.Unknown);

                Assert.True(section.Paragraphs[0].Text == "Section 5");
                Assert.True(section2.Paragraphs[0].Text == "Section 6");

                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins2.docx"))) {

                Assert.True(document.Sections[0].Margins.Type == WordMargin.Normal);
                Assert.True(document.Sections[1].Margins.Type == WordMargin.Narrow);
                Assert.True(document.Sections[2].Margins.Type == WordMargin.Mirrored);
                Assert.True(document.Sections[3].Margins.Type == WordMargin.Moderate);
                Assert.True(document.Sections[4].Margins.Type == WordMargin.Wide);
                Assert.True(document.Sections[5].Margins.Type == WordMargin.Office2003Default);
                Assert.True(document.Sections[6].Margins.Type == WordMargin.Unknown);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMargins2.docx"))) {

                Assert.True(document.Sections[0].Margins.Type == WordMargin.Normal);
                Assert.True(document.Sections[1].Margins.Type == WordMargin.Narrow);
                Assert.True(document.Sections[2].Margins.Type == WordMargin.Mirrored);
                Assert.True(document.Sections[3].Margins.Type == WordMargin.Moderate);
                Assert.True(document.Sections[4].Margins.Type == WordMargin.Wide);
                Assert.True(document.Sections[5].Margins.Type == WordMargin.Office2003Default);
                Assert.True(document.Sections[6].Margins.Type == WordMargin.Unknown);

            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithPageMarginsCentimeters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageMarginsCentimeters.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].Margins.BottomCentimeters = 2.30;
                document.Sections[0].Margins.TopCentimeters = 5.50;
                document.Sections[0].Margins.LeftCentimeters = 3.01;
                document.Sections[0].Margins.RightCentimeters = 3.05;

                Assert.True(document.Sections[0].Margins.BottomCentimeters == 2.2998236331569664);
                Assert.True(document.Sections[0].Margins.TopCentimeters == 5.499118165784832);
                Assert.True(document.Sections[0].Margins.LeftCentimeters == 3.0088183421516757);
                Assert.True(document.Sections[0].Margins.RightCentimeters == 3.049382716049383);

                Assert.True(document.Sections[0].Margins.Bottom == 1304);
                Assert.True(document.Sections[0].Margins.Top == 3118);
                Assert.True(document.Sections[0].Margins.Left.Value == 1706);
                Assert.True(document.Sections[0].Margins.Right.Value == 1729);



                document.Save(false);

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

        }
    }

}
