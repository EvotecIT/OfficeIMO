using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithDefaultHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddHeadersAndFooters();

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeader = document.Header.Default.AddParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                Assert.True(document.Header.Default.Paragraphs[0].Text == "Default Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                
                
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Default Header / Section 0", "Text for default header is wrong (section 0) (load)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during read is wrong (load). Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during read is wrong (load). Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during read is wrong. (load)");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong (load). Current: " + document.Sections[0].Paragraphs.Count);
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentWithDefaultHeadersAndFootersOddEven() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeaderO = document.Header.Default.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                Assert.True(document.Header.Default.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Header.Even.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Even Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault2.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Header.Default.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Header.Even.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Even Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithDefaultHeadersAndFootersOddEvenFirst() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefaultFirst1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeaderF = document.Header.First.AddParagraph();
                paragraphInHeaderF.Text = "First Header / Section 0";

                var paragraphInHeaderO = document.Header.Default.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                Assert.True(document.Header.Default.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Header.Even.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Even Header / Section 0", "Text for even header is wrong (section 0)");

                Assert.True(document.Header.First.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.First.Paragraphs[0].Text == "First Header / Section 0", "Text for first header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefaultFirst1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(document.Header.Default.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Default.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Header.Even.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.Even.Paragraphs[0].Text == "Even Header / Section 0", "Text for even header is wrong (section 0)");

                Assert.True(document.Header.First.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(document.Header.First.Paragraphs[0].Text == "First Header / Section 0", "Text for first header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");
            }
        }
        
        [Fact]
        public void Test_CreatingWordDocumentHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeaders.docx");
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
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeaders.docx"))) {

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
            }
        }
        [Fact]
        public void Test_CreatingWordDocumentHeadersAndFootersAndDeletingHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndRemoveThem.docx");
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


                var section1 = document.AddSection();
                section1.AddHeadersAndFooters();
                section1.DifferentFirstPage = true;
                section1.DifferentOddAndEvenPages = true;

                section1.Header.Default.AddParagraph().SetText("Test Section 0 - Header");
                section1.Footer.Default.AddParagraph().SetText("Test Section 0 - Footer");
                section1.Header.First.AddParagraph().SetText("Test Section 0 - First Header");
                section1.Footer.First.AddParagraph().SetText("Test Section 0 - First Footer");
                section1.Header.Even.AddParagraph().SetText("Test Section 0 - Header Even");
                section1.Footer.Even.AddParagraph().SetText("Test Section 0 - Footer Even");


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
                Assert.True(document.Sections.Count == 2, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.Save(false);
            }

            // One liner fixes
            WordHelpers.RemoveHeadersAndFooters(filePath);

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndRemoveThem.docx"))) {

                Assert.True(document.Header.Default == null);
                Assert.True(document.Footer.Default == null);
                Assert.True(document.Header.First == null);
                Assert.True(document.Footer.First== null);
                Assert.True(document.Header.Even== null);
                Assert.True(document.Footer.Even == null );

                Assert.True(document.Sections[1].Header.Default == null);
                Assert.True(document.Sections[1].Footer.Default == null);
                Assert.True(document.Sections[1].Header.First == null);
                Assert.True(document.Sections[1].Footer.First == null);
                Assert.True(document.Sections[1].Header.Even == null);
                Assert.True(document.Sections[1].Footer.Even == null);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 2, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
            }
        }

    }
}
