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
                Assert.True(document.Sections[0].PageOrientation == OfficePageOrientation.Portrait, "Page orientation should match");
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
                section2.Margins.Bottom = global::OfficeIMO.OfficeLength.Points(0.75D);
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

                document.Save();

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

                document.Sections[0].Margins.Bottom = global::OfficeIMO.OfficeLength.Centimeters(2.3D);
                document.Sections[0].Margins.Top = global::OfficeIMO.OfficeLength.Centimeters(5.5D);
                document.Sections[0].Margins.Left = global::OfficeIMO.OfficeLength.Centimeters(3.01D);
                document.Sections[0].Margins.Right = global::OfficeIMO.OfficeLength.Centimeters(3.05D);

                Assert.InRange(Math.Abs(document.Sections[0].Margins.Bottom!.Value.ToCentimeters() - 2.3), 0, 0.001);
                Assert.InRange(Math.Abs(document.Sections[0].Margins.Top!.Value.ToCentimeters() - 5.5), 0, 0.001);
                Assert.InRange(Math.Abs(document.Sections[0].Margins.Left!.Value.ToCentimeters() - 3.01), 0, 0.001);
                Assert.InRange(Math.Abs(document.Sections[0].Margins.Right!.Value.ToCentimeters() - 3.05), 0, 0.001);

                Assert.True(checked((int?)(document.Sections[0].Margins.Bottom?.ToTwips())) == 1304);
                Assert.True(checked((int?)(document.Sections[0].Margins.Top?.ToTwips())) == 3118);
                Assert.True(checked((uint)(document.Sections[0].Margins.Left.GetValueOrDefault().ToTwips())) == 1706);
                Assert.True(checked((uint)(document.Sections[0].Margins.Right.GetValueOrDefault().ToTwips())) == 1729);



                document.Save();

                Assert.True(HasUnexpectedElements(document) == false, "Document has unexpected elements. Order of elements matters!");
            }

        }

        [Fact]
        public void Test_DefaultMarginsWhenAddingSection() {
            using var document = WordDocument.Create();

            Assert.Equal(WordMargin.Normal, document.Sections[0].Margins.Type);
            Assert.Equal<UInt32>(1440U, checked((uint)(document.Sections[0].Margins.Left.GetValueOrDefault().ToTwips())));
            Assert.Equal<UInt32>(1440U, checked((uint)(document.Sections[0].Margins.Right.GetValueOrDefault().ToTwips())));
            Assert.Equal(1440, checked((int?)(document.Sections[0].Margins.Top?.ToTwips())));
            Assert.Equal(1440, checked((int?)(document.Sections[0].Margins.Bottom?.ToTwips())));

            var section = document.AddSection();

            Assert.Equal(WordMargin.Normal, section.Margins.Type);
            Assert.Equal<UInt32>(1440U, checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())));
            Assert.Equal<UInt32>(1440U, checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips())));
            Assert.Equal(1440, checked((int?)(section.Margins.Top?.ToTwips())));
            Assert.Equal(1440, checked((int?)(section.Margins.Bottom?.ToTwips())));
        }
    }

}
