using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Helper;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithSectionHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                document.Sections[0].Header.Default.AddParagraph().SetText("Test Section 0 - Header");




                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                section1.Header.Default.AddParagraph().SetText("Test Section 1 - Header");
                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);

                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;


                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(document.Sections[0].Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(document.Sections[1].Header.Default.Paragraphs[0].Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");
                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(document.Sections[0].Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(document.Sections[1].Header.Default.Paragraphs[0].Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");


                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during read is wrong (load). Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during read is wrong (load). Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during read is wrong. (load)");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong (load). Current: " + document.Sections[0].Paragraphs.Count);
            }
        }
    }
}
