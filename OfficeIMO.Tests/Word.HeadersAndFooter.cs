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
        public void Test_CreatingWordDocumentWithDefaultHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddHeadersAndFooters();

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeader = document.Header.Default.InsertParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                Assert.True(document.Header.Default.Paragraphs[0].Text == "Default Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                
                
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault.docx"))) {
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
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeaderO = document.Header.Default.InsertParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.InsertParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
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
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault.docx"))) {
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
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefaultFirst.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.DifferentOddAndEvenPages = true;
                document.DifferentFirstPage = true;

                var paragraph = document.InsertParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                var paragraphInHeaderF = document.Header.First.InsertParagraph();
                paragraphInHeaderF.Text = "First Header / Section 0";

                var paragraphInHeaderO = document.Header.Default.InsertParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var paragraphInHeaderE = document.Header.Even.InsertParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = System.Drawing.Color.Red.ToHexColor();

                document.InsertPageBreak();

                paragraph = document.InsertParagraph("Basic paragraph - Page 3");
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
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefaultFirst.docx"))) {
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
    }
}
