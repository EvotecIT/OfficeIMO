using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithPageSize() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageSize.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                Assert.True(document.Sections[0].PageOrientation == PageOrientationValues.Portrait, "Page orientation should match");
                Assert.True(document.Paragraphs.Count == 0, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");

                document.AddParagraph("Section 0");
                document.Sections[0].PageSettings.PageSize = WordPageSize.A3;

                document.AddSection();
                document.Sections[1].PageSettings.PageSize = WordPageSize.A4;
                document.AddParagraph("Section 1");

                document.AddSection();
                document.Sections[2].PageSettings.PageSize = WordPageSize.A5;
                document.AddParagraph("Section 2");

                document.AddSection();
                document.Sections[3].PageSettings.PageSize = WordPageSize.A6;
                document.AddParagraph("Section 3");

                document.AddSection();
                document.Sections[4].PageSettings.PageSize = WordPageSize.Executive;
                document.AddParagraph("Section 4");

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


                Assert.True(document.Sections[0].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[1].PageSettings.PageSize == WordPageSize.A4);
                Assert.True(document.Sections[2].PageSettings.PageSize == WordPageSize.A5);
                Assert.True(document.Sections[3].PageSettings.PageSize == WordPageSize.A6);
                Assert.True(document.Sections[4].PageSettings.PageSize == WordPageSize.Executive);
                Assert.True(document.Sections[5].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[6].PageSettings.PageSize == WordPageSize.Unknown);
                Assert.True(document.Sections[7].PageSettings.PageSize == WordPageSize.B5);
                Assert.True(document.Sections[8].PageSettings.PageSize == WordPageSize.Letter);
                Assert.True(document.Sections[9].PageSettings.PageSize == WordPageSize.Legal);


                Assert.True(section.PageSettings.PageSize == WordPageSize.A3);
                Assert.True(section2.PageSettings.PageSize == WordPageSize.Unknown);
                Assert.True(section3.PageSettings.PageSize == WordPageSize.B5);
                Assert.True(section4.PageSettings.PageSize == WordPageSize.Letter);
                Assert.True(section5.PageSettings.PageSize == WordPageSize.Legal);

                Assert.True(section.Paragraphs[0].Text == "Section 5");
                Assert.True(section2.Paragraphs[0].Text == "Section 6");

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageSize.docx"))) {
                Assert.True(document.Sections[0].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[1].PageSettings.PageSize == WordPageSize.A4);
                Assert.True(document.Sections[2].PageSettings.PageSize == WordPageSize.A5);
                Assert.True(document.Sections[3].PageSettings.PageSize == WordPageSize.A6);
                Assert.True(document.Sections[4].PageSettings.PageSize == WordPageSize.Executive);
                Assert.True(document.Sections[5].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[6].PageSettings.PageSize == WordPageSize.Unknown);
                Assert.True(document.Sections[7].PageSettings.PageSize == WordPageSize.B5);
                Assert.True(document.Sections[8].PageSettings.PageSize == WordPageSize.Letter);
                Assert.True(document.Sections[9].PageSettings.PageSize == WordPageSize.Legal);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithSectionsPageSize.docx"))) {
                Assert.True(document.Sections[0].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[1].PageSettings.PageSize == WordPageSize.A4);
                Assert.True(document.Sections[2].PageSettings.PageSize == WordPageSize.A5);
                Assert.True(document.Sections[3].PageSettings.PageSize == WordPageSize.A6);
                Assert.True(document.Sections[4].PageSettings.PageSize == WordPageSize.Executive);
                Assert.True(document.Sections[5].PageSettings.PageSize == WordPageSize.A3);
                Assert.True(document.Sections[6].PageSettings.PageSize == WordPageSize.Unknown);
                Assert.True(document.Sections[7].PageSettings.PageSize == WordPageSize.B5);
                Assert.True(document.Sections[8].PageSettings.PageSize == WordPageSize.Letter);
                Assert.True(document.Sections[9].PageSettings.PageSize == WordPageSize.Legal);
            }
        }
    }
}
