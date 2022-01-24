using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
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
        
        [Fact]
        public void Test_CreatingWordDocumentWithSectionHeadersAndFootersAdvanced() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                document.DifferentFirstPage = true;

                Assert.True(document.Header.First.Paragraphs.Count == 0, "First paragraph should not be there");
                document.Sections[0].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header", "First Header Should be correct");
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
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header", "First Header Should be correct");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(document.Sections[0].Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(document.Sections[1].Header.Default.Paragraphs[0].Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");

                document.Sections[0].Header.First.Paragraphs[0].Text = "Test Section 0 - First Header After mods";

                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during read is wrong (load). Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during read is wrong (load). Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during read is wrong. (load)");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong (load). Current: " + document.Sections[0].Paragraphs.Count);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header After mods", "First Header Should be correct");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(document.Sections[0].Header.Default.Paragraphs[0].Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(document.Sections[1].Header.Default.Paragraphs[0].Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");


                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during read is wrong (load). Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during read is wrong (load). Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during read is wrong. (load)");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong (load). Current: " + document.Sections[0].Paragraphs.Count);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentHeadersAndFootersWithSections() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();

                document.Header.Default.AddParagraph().SetText("Test Section 0 - Header");
                document.Footer.Default.AddParagraph().SetText("Test Section 0 - Footer");

                Assert.True(document.Header.First == null);
                Assert.True(document.Footer.First == null);

                document.DifferentFirstPage = true;

                Assert.True(document.Header.First != null);
                Assert.True(document.Footer.First != null);
                document.Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                document.Footer.First.AddParagraph().SetText("Test Section 0 - First Footer");

                Assert.True(document.Header.Even == null);
                Assert.True(document.Footer.Even == null);

                document.DifferentOddAndEvenPages = true;

                Assert.True(document.Header.Even != null);
                Assert.True(document.Footer.Even != null);

                document.Header.Even.AddParagraph().SetText("Test Section 0 - Header Even");
                document.Footer.Even.AddParagraph().SetText("Test Section 0 - Footer Even");

                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(document.Footer.Default.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(document.Footer.First.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(document.Footer.Even.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(document.Header.Default.Paragraphs.Count == 1);
                Assert.True(document.Footer.Default.Paragraphs.Count == 1);
                Assert.True(document.Header.First.Paragraphs.Count == 1);
                Assert.True(document.Footer.First.Paragraphs.Count == 1);
                Assert.True(document.Header.Even.Paragraphs.Count == 1);
                Assert.True(document.Footer.Even.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {

                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(document.Footer.Default.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(document.Footer.First.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(document.Footer.Even.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(document.Header.Default.Paragraphs.Count == 1);
                Assert.True(document.Footer.Default.Paragraphs.Count == 1);
                Assert.True(document.Header.First.Paragraphs.Count == 1);
                Assert.True(document.Footer.First.Paragraphs.Count == 1);
                Assert.True(document.Header.Even.Paragraphs.Count == 1);
                Assert.True(document.Footer.Even.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {

                Assert.True(document.Header.Default.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(document.Footer.Default.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(document.Header.First.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(document.Footer.First.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(document.Footer.Even.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(document.Header.Default.Paragraphs.Count == 1);
                Assert.True(document.Footer.Default.Paragraphs.Count == 1);
                Assert.True(document.Header.First.Paragraphs.Count == 1);
                Assert.True(document.Footer.First.Paragraphs.Count == 1);
                Assert.True(document.Header.Even.Paragraphs.Count == 1);
                Assert.True(document.Footer.Even.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.AddSection();
                document.Sections[1].PageOrientation = PageOrientationValues.Landscape;

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 2, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[1].Header.Default == null);
                Assert.True(document.Sections[1].Footer.Default == null);
                Assert.True(document.Sections[1].Header.First == null);
                Assert.True(document.Sections[1].Footer.First == null);
                Assert.True(document.Sections[1].Header.Even == null);
                Assert.True(document.Sections[1].Footer.Even == null);
                
                document.AddSection();
                document.Sections[2].PageOrientation = PageOrientationValues.Landscape;

                Assert.True(document.Sections[2].Header.Default == null);
                Assert.True(document.Sections[2].Footer.Default == null);
                Assert.True(document.Sections[2].Header.First == null);
                Assert.True(document.Sections[2].Footer.First == null);
                Assert.True(document.Sections[2].Header.Even == null);
                Assert.True(document.Sections[2].Footer.Even == null);

                document.Sections[2].AddHeadersAndFooters();

                Assert.True(document.Sections[2].Header.Default != null);
                Assert.True(document.Sections[2].Footer.Default != null);
                Assert.True(document.Sections[2].Header.First == null);
                Assert.True(document.Sections[2].Footer.First == null);
                Assert.True(document.Sections[2].Header.Even == null);
                Assert.True(document.Sections[2].Footer.Even == null);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");

                document.Sections[2].Header.Default.AddParagraph().SetText("Test Section 0 - Header");

                Assert.True(document.Sections[2].Header.Default.Paragraphs.Count == 1);

                document.AddSection();
                document.Sections[3].AddHeadersAndFooters();
                document.Sections[3].PageOrientation = PageOrientationValues.Landscape;

                document.Sections[1].AddHeadersAndFooters();
                document.Sections[1].DifferentOddAndEvenPages = true;

                document.Sections[1].Footer.Even.AddParagraph().SetText("Test Section 1 - Even");

                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                Assert.True(document.Sections[3].Header.Default != null);
                Assert.True(document.Sections[3].Footer.Default != null);
                Assert.True(document.Sections[3].Header.First == null);
                Assert.True(document.Sections[3].Footer.First == null);
                Assert.True(document.Sections[3].Header.Even == null);
                Assert.True(document.Sections[3].Footer.Even == null);

                document.Sections[3].DifferentFirstPage = true;
                document.Sections[3].DifferentOddAndEvenPages = true;

                document.Sections[3].Header.Default.AddParagraph().SetText("Test Section 0 - Header");
                document.Sections[3].Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                document.Sections[3].Header.Even.AddParagraph().SetText("Test Section 0 - Even");

                Assert.True(document.Sections[3].Header.Default != null);
                Assert.True(document.Sections[3].Footer.Default != null);
                Assert.True(document.Sections[3].Header.First != null);
                Assert.True(document.Sections[3].Footer.First != null);
                Assert.True(document.Sections[3].Header.Even != null);
                Assert.True(document.Sections[3].Footer.Even != null);


                Assert.True(document.Sections[2].Header.Default.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(document.Sections[1].Footer.Even.Paragraphs[0].Text == "Test Section 1 - Even");


                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                document.Save();
            }
        }

    }
}
