using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for headers and footers.
    /// </summary>
    public partial class Word {
        private static WordHeader RequireSectionHeader(WordDocument doc, int index, HeaderFooterValues type) {
            Assert.NotNull(doc);
            Assert.InRange(index, 0, doc.Sections.Count - 1);

            var section = doc.Sections[index];
            if (type == HeaderFooterValues.Default && section.Header.Default == null) {
                section.AddHeadersAndFooters();
            }

            if (type == HeaderFooterValues.Default) {
                return Assert.IsAssignableFrom<WordHeader>(section.Header.Default);
            }

            if (type == HeaderFooterValues.Even) {
                return Assert.IsAssignableFrom<WordHeader>(section.Header.Even);
            }

            if (type == HeaderFooterValues.First) {
                return Assert.IsAssignableFrom<WordHeader>(section.Header.First);
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported header type.");
        }

        private static WordFooter RequireSectionFooter(WordDocument doc, int index, HeaderFooterValues type) {
            Assert.NotNull(doc);
            Assert.InRange(index, 0, doc.Sections.Count - 1);

            var section = doc.Sections[index];
            if (type == HeaderFooterValues.Default && section.Footer.Default == null) {
                section.AddHeadersAndFooters();
            }

            if (type == HeaderFooterValues.Default) {
                return Assert.IsAssignableFrom<WordFooter>(section.Footer.Default);
            }

            if (type == HeaderFooterValues.Even) {
                return Assert.IsAssignableFrom<WordFooter>(section.Footer.Even);
            }

            if (type == HeaderFooterValues.First) {
                return Assert.IsAssignableFrom<WordFooter>(section.Footer.First);
            }

            throw new ArgumentOutOfRangeException(nameof(type), type, "Unsupported footer type.");
        }

        [Fact]
        public void Test_CreatingWordDocumentWithDefaultHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.AddHeadersAndFooters();

                var paragraph = document.AddParagraph("Basic paragraph - Page 1");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var paragraphInHeader = defaultHeader.AddParagraph();
                paragraphInHeader.Text = "Default Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                Assert.True(defaultHeader.Paragraphs[0].Text == "Default Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault1.docx"))) {
                // There is only one Paragraph at the document level.
                var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                Assert.True(defaultHeader.Paragraphs[0].Text == "Default Header / Section 0", "Text for default header is wrong (section 0) (load)");

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
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var oddHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var paragraphInHeaderO = oddHeader.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                var paragraphInHeaderE = evenHeader.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                Assert.True(oddHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(oddHeader.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(evenHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(evenHeader.Paragraphs[0].Text == "Even Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefault2.docx"))) {
                // There is only one Paragraph at the document level.
                var oddHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                Assert.True(oddHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(oddHeader.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                Assert.True(evenHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(evenHeader.Paragraphs[0].Text == "Even Header / Section 0", "Text for default header is wrong (section 0)");

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
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                var firstHeader = RequireSectionHeader(document, 0, HeaderFooterValues.First);
                var paragraphInHeaderF = firstHeader.AddParagraph();
                paragraphInHeaderF.Text = "First Header / Section 0";

                var oddHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var paragraphInHeaderO = oddHeader.AddParagraph();
                paragraphInHeaderO.Text = "Odd Header / Section 0";

                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                var paragraphInHeaderE = evenHeader.AddParagraph();
                paragraphInHeaderE.Text = "Even Header / Section 0";

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 2");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                document.AddPageBreak();

                paragraph = document.AddParagraph("Basic paragraph - Page 3");
                paragraph.ParagraphAlignment = JustificationValues.Center;
                paragraph.Color = SixLabors.ImageSharp.Color.Red;

                Assert.True(oddHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(oddHeader.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                Assert.True(evenHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(evenHeader.Paragraphs[0].Text == "Even Header / Section 0", "Text for even header is wrong (section 0)");

                Assert.True(firstHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(firstHeader.Paragraphs[0].Text == "First Header / Section 0", "Text for first header is wrong (section 0)");

                Assert.True(document.Paragraphs.Count == 5, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 2, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 5, "Number of paragraphs on 1st section is wrong.");


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersDefaultFirst1.docx"))) {
                // There is only one Paragraph at the document level.
                var oddHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                Assert.True(oddHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(oddHeader.Paragraphs[0].Text == "Odd Header / Section 0", "Text for default header is wrong (section 0)");

                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                Assert.True(evenHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(evenHeader.Paragraphs[0].Text == "Even Header / Section 0", "Text for even header is wrong (section 0)");

                var firstHeader = RequireSectionHeader(document, 0, HeaderFooterValues.First);
                Assert.True(firstHeader.Paragraphs.Count == 1, "Should only have X paragraphs");
                Assert.True(firstHeader.Paragraphs[0].Text == "First Header / Section 0", "Text for first header is wrong (section 0)");

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

                var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);

                defaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                defaultFooter.AddParagraph().SetText("Test Section 0 - Footer");

                Assert.True(document.Sections[0].Header.First == null);
                Assert.True(document.Sections[0].Footer.First == null);

                document.DifferentFirstPage = true;

                Assert.True(document.Sections[0].Header.First != null);
                Assert.True(document.Sections[0].Footer.First != null);

                var firstHeader = RequireSectionHeader(document, 0, HeaderFooterValues.First);
                var firstFooter = RequireSectionFooter(document, 0, HeaderFooterValues.First);

                firstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                firstFooter.AddParagraph().SetText("Test Section 0 - First Footer");

                Assert.True(document.Sections[0].Header.Even == null);
                Assert.True(document.Sections[0].Footer.Even == null);

                document.DifferentOddAndEvenPages = true;

                Assert.True(document.Sections[0].Header.Even != null);
                Assert.True(document.Sections[0].Footer.Even != null);

                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                var evenFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Even);

                evenHeader.AddParagraph().SetText("Test Section 0 - Header Even");
                evenFooter.AddParagraph().SetText("Test Section 0 - Footer Even");

                Assert.True(defaultHeader.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(defaultFooter.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(firstHeader.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(firstFooter.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(evenHeader.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(evenFooter.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(defaultHeader.Paragraphs.Count == 1);
                Assert.True(defaultFooter.Paragraphs.Count == 1);
                Assert.True(firstHeader.Paragraphs.Count == 1);
                Assert.True(firstFooter.Paragraphs.Count == 1);
                Assert.True(evenHeader.Paragraphs.Count == 1);
                Assert.True(evenFooter.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeaders.docx"))) {

                var defaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var defaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var firstHeader = RequireSectionHeader(document, 0, HeaderFooterValues.First);
                var firstFooter = RequireSectionFooter(document, 0, HeaderFooterValues.First);
                var evenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                var evenFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Even);

                Assert.True(defaultHeader.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(defaultFooter.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(firstHeader.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(firstFooter.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(evenHeader.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(evenFooter.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(defaultHeader.Paragraphs.Count == 1);
                Assert.True(defaultFooter.Paragraphs.Count == 1);
                Assert.True(firstHeader.Paragraphs.Count == 1);
                Assert.True(firstFooter.Paragraphs.Count == 1);
                Assert.True(evenHeader.Paragraphs.Count == 1);
                Assert.True(evenFooter.Paragraphs.Count == 1);

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

                var section0DefaultHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var section0DefaultFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);

                section0DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                section0DefaultFooter.AddParagraph().SetText("Test Section 0 - Footer");

                Assert.True(document.Header!.First == null);
                Assert.True(document.Footer!.First == null);

                document.DifferentFirstPage = true;

                var section0FirstHeader = RequireSectionHeader(document, 0, HeaderFooterValues.First);
                var section0FirstFooter = RequireSectionFooter(document, 0, HeaderFooterValues.First);

                Assert.NotNull(section0FirstHeader);
                Assert.NotNull(section0FirstFooter);
                section0FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                section0FirstFooter.AddParagraph().SetText("Test Section 0 - First Footer");

                Assert.True(document.Header!.Even == null);
                Assert.True(document.Footer!.Even == null);

                document.DifferentOddAndEvenPages = true;

                var section0EvenHeader = RequireSectionHeader(document, 0, HeaderFooterValues.Even);
                var section0EvenFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Even);

                Assert.NotNull(section0EvenHeader);
                Assert.NotNull(section0EvenFooter);
                section0EvenHeader.AddParagraph().SetText("Test Section 0 - Header Even");
                section0EvenFooter.AddParagraph().SetText("Test Section 0 - Footer Even");


                var section1 = document.AddSection();
                section1.AddHeadersAndFooters();
                section1.DifferentFirstPage = true;
                section1.DifferentOddAndEvenPages = true;

                var section1DefaultHeader = RequireSectionHeader(document, 1, HeaderFooterValues.Default);
                var section1DefaultFooter = RequireSectionFooter(document, 1, HeaderFooterValues.Default);
                section1DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                section1DefaultFooter.AddParagraph().SetText("Test Section 0 - Footer");

                var section1FirstHeader = RequireSectionHeader(document, 1, HeaderFooterValues.First);
                var section1FirstFooter = RequireSectionFooter(document, 1, HeaderFooterValues.First);
                section1FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                section1FirstFooter.AddParagraph().SetText("Test Section 0 - First Footer");

                var section1EvenHeader = RequireSectionHeader(document, 1, HeaderFooterValues.Even);
                var section1EvenFooter = RequireSectionFooter(document, 1, HeaderFooterValues.Even);
                section1EvenHeader.AddParagraph().SetText("Test Section 0 - Header Even");
                section1EvenFooter.AddParagraph().SetText("Test Section 0 - Footer Even");


                Assert.True(section0DefaultHeader.Paragraphs[0].Text == "Test Section 0 - Header");
                Assert.True(section0DefaultFooter.Paragraphs[0].Text == "Test Section 0 - Footer");
                Assert.True(section0FirstHeader.Paragraphs[0].Text == "Test Section 0 - First Header");
                Assert.True(section0FirstFooter.Paragraphs[0].Text == "Test Section 0 - First Footer");
                Assert.True(section0EvenHeader.Paragraphs[0].Text == "Test Section 0 - Header Even");
                Assert.True(section0EvenFooter.Paragraphs[0].Text == "Test Section 0 - Footer Even");

                Assert.True(section0DefaultHeader.Paragraphs.Count == 1);
                Assert.True(section0DefaultFooter.Paragraphs.Count == 1);
                Assert.True(section0FirstHeader.Paragraphs.Count == 1);
                Assert.True(section0FirstFooter.Paragraphs.Count == 1);
                Assert.True(section0EvenHeader.Paragraphs.Count == 1);
                Assert.True(section0EvenFooter.Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 2, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.Save(false);
            }

            // One liner fixes
            WordHelpers.RemoveHeadersAndFooters(filePath);

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndRemoveThem.docx"))) {

                var section0Headers = document.Header;
                Assert.NotNull(section0Headers);
                Assert.Null(section0Headers.Default);
                Assert.Null(section0Headers.First);
                Assert.Null(section0Headers.Even);

                var section0Footers = document.Footer;
                Assert.NotNull(section0Footers);
                Assert.Null(section0Footers.Default);
                Assert.Null(section0Footers.First);
                Assert.Null(section0Footers.Even);

                var section1Headers = document.Sections[1].Header;
                Assert.NotNull(section1Headers);
                Assert.Null(section1Headers.Default);
                Assert.Null(section1Headers.First);
                Assert.Null(section1Headers.Even);

                var section1Footers = document.Sections[1].Footer;
                Assert.NotNull(section1Footers);
                Assert.Null(section1Footers.Default);
                Assert.Null(section1Footers.First);
                Assert.Null(section1Footers.Even);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 2, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
            }
        }

    }
}
