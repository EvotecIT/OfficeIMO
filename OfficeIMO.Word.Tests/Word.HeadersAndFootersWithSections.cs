using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {

        private static WordHeaders GetDocumentHeaders(WordDocument document, string context) {
            Assert.NotNull(document.Header);
            return document.Header ?? throw new InvalidOperationException("Headers were not created for {context}.");
        }

        private static WordFooters GetDocumentFooters(WordDocument document, string context) {
            Assert.NotNull(document.Footer);
            return document.Footer ?? throw new InvalidOperationException("Footers were not created for {context}.");
        }

        private static WordHeaders GetSectionHeaders(WordDocument document, int sectionIndex, string context) {
            Assert.True(sectionIndex >= 0, $"Section index must be non-negative for {context}.");
            Assert.True(document.Sections.Count > sectionIndex, $"Section index {sectionIndex} is out of range for {context}.");
            var headers = document.Sections[sectionIndex].Header;
            Assert.NotNull(headers);
            return headers ?? throw new InvalidOperationException("Section headers were not created for {context}.");
        }

        private static WordFooters GetSectionFooters(WordDocument document, int sectionIndex, string context) {
            Assert.True(sectionIndex >= 0, $"Section index must be non-negative for {context}.");
            Assert.True(document.Sections.Count > sectionIndex, $"Section index {sectionIndex} is out of range for {context}.");
            var footers = document.Sections[sectionIndex].Footer;
            Assert.NotNull(footers);
            return footers ?? throw new InvalidOperationException("Section footers were not created for {context}.");
        }

        private static WordHeader GetDefaultHeader(WordHeaders headers, string context) {
            Assert.NotNull(headers.Default);
            return headers.Default ?? throw new InvalidOperationException("Default header was not created for {context}.");
        }

        private static WordHeader GetFirstHeader(WordHeaders headers, string context) {
            Assert.NotNull(headers.First);
            return headers.First ?? throw new InvalidOperationException("First header was not created for {context}.");
        }

        private static WordHeader GetEvenHeader(WordHeaders headers, string context) {
            Assert.NotNull(headers.Even);
            return headers.Even ?? throw new InvalidOperationException("Even header was not created for {context}.");
        }

        private static WordFooter GetDefaultFooter(WordFooters footers, string context) {
            Assert.NotNull(footers.Default);
            return footers.Default ?? throw new InvalidOperationException("Default footer was not created for {context}.");
        }

        private static WordFooter GetFirstFooter(WordFooters footers, string context) {
            Assert.NotNull(footers.First);
            return footers.First ?? throw new InvalidOperationException("First footer was not created for {context}.");
        }

        private static WordFooter GetEvenFooter(WordFooters footers, string context) {
            Assert.NotNull(footers.Even);
            return footers.Even ?? throw new InvalidOperationException("Even footer was not created for {context}.");
        }

        private static WordParagraph GetParagraphAt(IList<WordParagraph> paragraphs, int index, string context) {
            Assert.NotEmpty(paragraphs);
            Assert.True(index >= 0 && index < paragraphs.Count, $"Paragraph index {index} is out of range for {context}.");
            return paragraphs[index];
        }

        private static WordHeader GetDocumentDefaultHeader(WordDocument document, string context) => GetDefaultHeader(GetDocumentHeaders(document, context), context);

        private static WordHeader GetDocumentFirstHeader(WordDocument document, string context) => GetFirstHeader(GetDocumentHeaders(document, context), context);

        private static WordHeader GetDocumentEvenHeader(WordDocument document, string context) => GetEvenHeader(GetDocumentHeaders(document, context), context);

        private static WordFooter GetDocumentDefaultFooter(WordDocument document, string context) => GetDefaultFooter(GetDocumentFooters(document, context), context);

        private static WordFooter GetDocumentFirstFooter(WordDocument document, string context) => GetFirstFooter(GetDocumentFooters(document, context), context);

        private static WordFooter GetDocumentEvenFooter(WordDocument document, string context) => GetEvenFooter(GetDocumentFooters(document, context), context);

        private static WordHeader GetSectionDefaultHeader(WordDocument document, int sectionIndex, string context) => GetDefaultHeader(GetSectionHeaders(document, sectionIndex, context), context);

        private static WordHeader GetSectionFirstHeader(WordDocument document, int sectionIndex, string context) => GetFirstHeader(GetSectionHeaders(document, sectionIndex, context), context);

        private static WordHeader GetSectionEvenHeader(WordDocument document, int sectionIndex, string context) => GetEvenHeader(GetSectionHeaders(document, sectionIndex, context), context);

        private static WordFooter GetSectionDefaultFooter(WordDocument document, int sectionIndex, string context) => GetDefaultFooter(GetSectionFooters(document, sectionIndex, context), context);

        private static WordFooter GetSectionFirstFooter(WordDocument document, int sectionIndex, string context) => GetFirstFooter(GetSectionFooters(document, sectionIndex, context), context);

        private static WordFooter GetSectionEvenFooter(WordDocument document, int sectionIndex, string context) => GetEvenFooter(GetSectionFooters(document, sectionIndex, context), context);

        private static WordHeaders GetSectionHeaders(WordSection section, string context) {
            Assert.NotNull(section);
            var headers = section.Header;
            Assert.NotNull(headers);
            return headers ?? throw new InvalidOperationException("Section headers were not created for {context}.");
        }

        private static WordFooters GetSectionFooters(WordSection section, string context) {
            Assert.NotNull(section);
            var footers = section.Footer;
            Assert.NotNull(footers);
            return footers ?? throw new InvalidOperationException("Section footers were not created for {context}.");
        }

        private static WordHeader GetSectionDefaultHeader(WordSection section, string context) => GetDefaultHeader(GetSectionHeaders(section, context), context);

        private static WordHeader GetSectionFirstHeader(WordSection section, string context) => GetFirstHeader(GetSectionHeaders(section, context), context);

        private static WordHeader GetSectionEvenHeader(WordSection section, string context) => GetEvenHeader(GetSectionHeaders(section, context), context);

        private static WordFooter GetSectionDefaultFooter(WordSection section, string context) => GetDefaultFooter(GetSectionFooters(section, context), context);

        private static WordFooter GetSectionFirstFooter(WordSection section, string context) => GetFirstFooter(GetSectionFooters(section, context), context);

        private static WordFooter GetSectionEvenFooter(WordSection section, string context) => GetEvenFooter(GetSectionFooters(section, context), context);

        [Fact]
        public void Test_CreatingWordDocumentWithSectionHeadersAndFooters() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {

                document.Sections[0].PageOrientation = PageOrientationValues.Landscape;
                document.AddParagraph("Test Section0");
                document.AddHeadersAndFooters();
                var section0Headers = GetSectionHeaders(document, 0, "section 0 during creation");
                GetDefaultHeader(section0Headers, "section 0 default header during creation").AddParagraph().SetText("Test Section 0 - Header");

                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                var section1Headers = GetSectionHeaders(section1, "section 1 during creation");
                GetDefaultHeader(section1Headers, "section 1 default header during creation").AddParagraph().SetText("Test Section 1 - Header");
                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);

                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;


                var documentDefaultHeader = GetDocumentDefaultHeader(document, "document default header during creation");
                Assert.True(GetParagraphAt(documentDefaultHeader.Paragraphs, 0, "document default header paragraphs during creation").Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                var section0DefaultHeader = GetDefaultHeader(section0Headers, "section 0 default header during creation");
                Assert.True(GetParagraphAt(section0DefaultHeader.Paragraphs, 0, "section 0 default header paragraphs during creation").Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                var section1DefaultHeader = GetDefaultHeader(section1Headers, "section 1 default header during creation");
                Assert.True(GetParagraphAt(section1DefaultHeader.Paragraphs, 0, "section 1 default header paragraphs during creation").Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");
                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx"))) {
                // There is only one Paragraph at the document level.
                var documentDefaultHeader = GetDocumentDefaultHeader(document, "document default header during initial load");
                Assert.True(GetParagraphAt(documentDefaultHeader.Paragraphs, 0, "document default header paragraphs during initial load").Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                var section0DefaultHeader = GetSectionDefaultHeader(document, 0, "section 0 default header during initial load");
                Assert.True(GetParagraphAt(section0DefaultHeader.Paragraphs, 0, "section 0 default header paragraphs during initial load").Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                var section1DefaultHeader = GetSectionDefaultHeader(document, 1, "section 1 default header during initial load");
                Assert.True(GetParagraphAt(section1DefaultHeader.Paragraphs, 0, "section 1 default header paragraphs during initial load").Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");


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

                Assert.True(GetDocumentFirstHeader(document, "document first header").Paragraphs.Count == 0, "First paragraph should not be there");
                GetSectionFirstHeader(document, 0, "section 0 first header").AddParagraph().SetText("Test Section 0 - First Header");
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header", "First Header Should be correct");
                GetSectionDefaultHeader(document, 0, "section 0 default header").AddParagraph().SetText("Test Section 0 - Header");

                var section1 = document.AddSection();
                section1.PageOrientation = PageOrientationValues.Portrait;
                section1.AddParagraph("Test Section1");
                section1.AddHeadersAndFooters();
                GetSectionDefaultHeader(section1, "section1 default header").AddParagraph().SetText("Test Section 1 - Header");
                //Console.WriteLine("Section 0 - Text 0: " + document.Sections[0].Paragraphs[0].Text);

                var section2 = document.AddSection();
                section2.AddParagraph("Test Section2");
                section2.PageOrientation = PageOrientationValues.Landscape;


                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 0, "section 0 default header").Paragraphs, 0, "section 0 default header paragraphs").Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 1, "section 1 default header").Paragraphs, 0, "section 1 default header paragraphs").Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");
                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);


                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header", "First Header Should be correct");
                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 0, "section 0 default header").Paragraphs, 0, "section 0 default header paragraphs").Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 1, "section 1 default header").Paragraphs, 0, "section 1 default header paragraphs").Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");

                GetParagraphAt(GetSectionFirstHeader(document, 0, "section 0 first header").Paragraphs, 0, "section 0 first header paragraphs").Text = "Test Section 0 - First Header After mods";

                Assert.True(document.Paragraphs.Count == 3, "Number of paragraphs during read is wrong (load). Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during read is wrong (load). Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during read is wrong. (load)");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong (load). Current: " + document.Sections[0].Paragraphs.Count);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndFootersSection1.docx"))) {
                // There is only one Paragraph at the document level.
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header After mods", "First Header Should be correct");
                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header", "Text for default header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 0, "section 0 default header").Paragraphs, 0, "section 0 default header paragraphs").Text == "Test Section 0 - Header", "Text for section header is wrong (section 0)");
                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 1, "section 1 default header").Paragraphs, 0, "section 1 default header paragraphs").Text == "Test Section 1 - Header", "Text for section header is wrong (section 1)");


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

                GetDocumentDefaultHeader(document, "document default header").AddParagraph().SetText("Test Section 0 - Header");
                GetDocumentDefaultFooter(document, "document default footer").AddParagraph().SetText("Test Section 0 - Footer");

                var documentHeaders = GetDocumentHeaders(document, "document headers state before first header creation");
                Assert.Null(documentHeaders.First);
                var documentFooters = GetDocumentFooters(document, "document footers state before first footer creation");
                Assert.Null(documentFooters.First);

                document.DifferentFirstPage = true;

                var documentFirstHeader = GetDocumentFirstHeader(document, "document first header after enabling DifferentFirstPage");
                var documentFirstFooter = GetDocumentFirstFooter(document, "document first footer after enabling DifferentFirstPage");
                Assert.NotNull(documentFirstHeader);
                Assert.NotNull(documentFirstFooter);
                documentFirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                documentFirstFooter.AddParagraph().SetText("Test Section 0 - First Footer");

                Assert.Null(documentHeaders.Even);
                Assert.Null(documentFooters.Even);

                document.DifferentOddAndEvenPages = true;

                var documentEvenHeader = GetDocumentEvenHeader(document, "document even header after enabling DifferentOddAndEvenPages");
                var documentEvenFooter = GetDocumentEvenFooter(document, "document even footer after enabling DifferentOddAndEvenPages");
                Assert.NotNull(documentEvenHeader);
                Assert.NotNull(documentEvenFooter);

                documentEvenHeader.AddParagraph().SetText("Test Section 0 - Header Even");
                documentEvenFooter.AddParagraph().SetText("Test Section 0 - Footer Even");

                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header");
                Assert.True(GetParagraphAt(GetDocumentDefaultFooter(document, "document default footer").Paragraphs, 0, "document default footer paragraphs").Text == "Test Section 0 - Footer");
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header");
                Assert.True(GetParagraphAt(GetDocumentFirstFooter(document, "document first footer").Paragraphs, 0, "document first footer paragraphs").Text == "Test Section 0 - First Footer");
                Assert.True(GetParagraphAt(GetDocumentEvenHeader(document, "document even header").Paragraphs, 0, "document even header paragraphs").Text == "Test Section 0 - Header Even");
                Assert.True(GetParagraphAt(GetDocumentEvenFooter(document, "document even footer").Paragraphs, 0, "document even footer paragraphs").Text == "Test Section 0 - Footer Even");

                Assert.True(GetDocumentDefaultHeader(document, "document default header").Paragraphs.Count == 1);
                Assert.True(GetDocumentDefaultFooter(document, "document default footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstHeader(document, "document first header").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstFooter(document, "document first footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenHeader(document, "document even header").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenFooter(document, "document even footer").Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {

                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header");
                Assert.True(GetParagraphAt(GetDocumentDefaultFooter(document, "document default footer").Paragraphs, 0, "document default footer paragraphs").Text == "Test Section 0 - Footer");
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header");
                Assert.True(GetParagraphAt(GetDocumentFirstFooter(document, "document first footer").Paragraphs, 0, "document first footer paragraphs").Text == "Test Section 0 - First Footer");
                Assert.True(GetParagraphAt(GetDocumentEvenHeader(document, "document even header").Paragraphs, 0, "document even header paragraphs").Text == "Test Section 0 - Header Even");
                Assert.True(GetParagraphAt(GetDocumentEvenFooter(document, "document even footer").Paragraphs, 0, "document even footer paragraphs").Text == "Test Section 0 - Footer Even");

                Assert.True(GetDocumentDefaultHeader(document, "document default header").Paragraphs.Count == 1);
                Assert.True(GetDocumentDefaultFooter(document, "document default footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstHeader(document, "document first header").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstFooter(document, "document first footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenHeader(document, "document even header").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenFooter(document, "document even footer").Paragraphs.Count == 1);

                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.PageBreaks.Count == 0, "Number of page breaks during creation is wrong. Current: " + document.PageBreaks.Count);
                Assert.True(document.Sections.Count == 1, "Number of sections during creation is wrong.");
                Assert.True(document.Sections[0].Paragraphs.Count == 1, "Number of paragraphs on 1st section is wrong. Current: " + document.Sections[0].Paragraphs.Count);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {

                Assert.True(GetParagraphAt(GetDocumentDefaultHeader(document, "document default header").Paragraphs, 0, "document default header paragraphs").Text == "Test Section 0 - Header");
                Assert.True(GetParagraphAt(GetDocumentDefaultFooter(document, "document default footer").Paragraphs, 0, "document default footer paragraphs").Text == "Test Section 0 - Footer");
                Assert.True(GetParagraphAt(GetDocumentFirstHeader(document, "document first header").Paragraphs, 0, "document first header paragraphs").Text == "Test Section 0 - First Header");
                Assert.True(GetParagraphAt(GetDocumentFirstFooter(document, "document first footer").Paragraphs, 0, "document first footer paragraphs").Text == "Test Section 0 - First Footer");
                Assert.True(GetParagraphAt(GetDocumentEvenHeader(document, "document even header").Paragraphs, 0, "document even header paragraphs").Text == "Test Section 0 - Header Even");
                Assert.True(GetParagraphAt(GetDocumentEvenFooter(document, "document even footer").Paragraphs, 0, "document even footer paragraphs").Text == "Test Section 0 - Footer Even");

                Assert.True(GetDocumentDefaultHeader(document, "document default header").Paragraphs.Count == 1);
                Assert.True(GetDocumentDefaultFooter(document, "document default footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstHeader(document, "document first header").Paragraphs.Count == 1);
                Assert.True(GetDocumentFirstFooter(document, "document first footer").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenHeader(document, "document even header").Paragraphs.Count == 1);
                Assert.True(GetDocumentEvenFooter(document, "document even footer").Paragraphs.Count == 1);

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

                var section1Headers = GetSectionHeaders(document, 1, "section 1 headers state after reload");
                Assert.Null(section1Headers.Default);
                Assert.Null(section1Headers.First);
                Assert.Null(section1Headers.Even);
                var section1Footers = GetSectionFooters(document, 1, "section 1 footers state after reload");
                Assert.Null(section1Footers.Default);
                Assert.Null(section1Footers.First);
                Assert.Null(section1Footers.Even);

                var section2 = document.AddSection();
                section2.PageOrientation = PageOrientationValues.Landscape;

                var section2Headers = GetSectionHeaders(section2, "section 2 headers state prior to AddHeadersAndFooters");
                Assert.Null(section2Headers.Default);
                Assert.Null(section2Headers.First);
                Assert.Null(section2Headers.Even);
                var section2Footers = GetSectionFooters(section2, "section 2 footers state prior to AddHeadersAndFooters");
                Assert.Null(section2Footers.Default);
                Assert.Null(section2Footers.First);
                Assert.Null(section2Footers.Even);

                section2.AddHeadersAndFooters();

                var section2DefaultHeader = GetSectionDefaultHeader(section2, "section 2 default header after AddHeadersAndFooters");
                var section2DefaultFooter = GetSectionDefaultFooter(section2, "section 2 default footer after AddHeadersAndFooters");
                Assert.NotNull(section2DefaultHeader);
                Assert.NotNull(section2DefaultFooter);
                Assert.Null(section2Headers.First);
                Assert.Null(section2Footers.First);
                Assert.Null(section2Headers.Even);
                Assert.Null(section2Footers.Even);

                document.Save();
            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 3, "Number of sections during creation is wrong.");

                GetSectionDefaultHeader(document, 2, "section 2 default header").AddParagraph().SetText("Test Section 0 - Header");

                Assert.True(GetSectionDefaultHeader(document, 2, "section 2 default header").Paragraphs.Count == 1);

                var newSection3 = document.AddSection();
                newSection3.AddHeadersAndFooters();
                newSection3.PageOrientation = PageOrientationValues.Landscape;

                var updatedSection1 = document.Sections[1];
                updatedSection1.AddHeadersAndFooters();
                updatedSection1.DifferentOddAndEvenPages = true;

                var section1EvenFooter = GetSectionEvenFooter(updatedSection1, "section 1 even footer after enabling DifferentOddAndEvenPages");
                Assert.NotNull(section1EvenFooter);
                section1EvenFooter.AddParagraph().SetText("Test Section 1 - Even");

                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSections.docx"))) {
                Assert.True(document.Paragraphs.Count == 1, "Number of paragraphs during creation is wrong. Current: " + document.Paragraphs.Count);
                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                var section3 = document.Sections[3];
                var section3Headers = GetSectionHeaders(section3, "section 3 headers state before enabling page differences");
                Assert.NotNull(section3Headers.Default);
                Assert.Null(section3Headers.First);
                Assert.Null(section3Headers.Even);
                var section3Footers = GetSectionFooters(section3, "section 3 footers state before enabling page differences");
                Assert.NotNull(section3Footers.Default);
                Assert.Null(section3Footers.First);
                Assert.Null(section3Footers.Even);

                section3.DifferentFirstPage = true;
                section3.DifferentOddAndEvenPages = true;

                var section3DefaultHeader = GetSectionDefaultHeader(section3, "section 3 default header after enabling page differences");
                var section3FirstHeader = GetSectionFirstHeader(section3, "section 3 first header after enabling page differences");
                var section3EvenHeader = GetSectionEvenHeader(section3, "section 3 even header after enabling page differences");
                var section3DefaultFooter = GetSectionDefaultFooter(section3, "section 3 default footer after enabling page differences");
                var section3FirstFooter = GetSectionFirstFooter(section3, "section 3 first footer after enabling page differences");
                var section3EvenFooter = GetSectionEvenFooter(section3, "section 3 even footer after enabling page differences");

                Assert.NotNull(section3DefaultHeader);
                Assert.NotNull(section3FirstHeader);
                Assert.NotNull(section3EvenHeader);
                Assert.NotNull(section3DefaultFooter);
                Assert.NotNull(section3FirstFooter);
                Assert.NotNull(section3EvenFooter);

                section3DefaultHeader.AddParagraph().SetText("Test Section 0 - Header");
                section3FirstHeader.AddParagraph().SetText("Test Section 0 - First Header");
                section3EvenHeader.AddParagraph().SetText("Test Section 0 - Even");


                Assert.True(GetParagraphAt(GetSectionDefaultHeader(document, 2, "section 2 default header").Paragraphs, 0, "section 2 default header paragraphs").Text == "Test Section 0 - Header");
                Assert.True(GetParagraphAt(GetSectionEvenFooter(document, 1, "section 1 even footer").Paragraphs, 0, "section 1 even footer paragraphs").Text == "Test Section 1 - Even");


                Assert.True(document.Sections.Count == 4, "Number of sections during creation is wrong.");

                document.Save();
            }
        }


        [Fact]
        public void Test_CreatingWordDocumentHeadersAndFootersOddEvenFirst() {
            string filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSectionsOddEventFirst.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Basic paragraph");

                document.AddHeadersAndFooters();
                GetSectionDefaultHeader(document, 0, "section 0 default header").AddParagraph("Test2").AddText("Section 0");

                var section1 = document.AddSection();
                section1.AddParagraph("Test Middle1 Section - 1");
                section1.AddHeadersAndFooters();
                GetSectionDefaultHeader(section1, "section1 default header").AddParagraph().AddText("Section 1 - Header");
                GetSectionDefaultFooter(section1, "section1 default footer").AddParagraph().AddText("Section 1 - Footer");

                var section2 = document.AddSection();
                section2.AddParagraph("Test Middle2 Section - 1");
                section2.AddHeadersAndFooters();
                GetSectionDefaultHeader(section2, "section2 default header").AddParagraph().AddText("Section 2 - Header");
                GetSectionDefaultFooter(section2, "section2 default footer").AddParagraph().AddText("Section 2 - Footer");

                var section3 = document.AddSection();
                section3.AddParagraph("Test Last Section - 1");
                section3.AddHeadersAndFooters();
                section3.DifferentOddAndEvenPages = true;
                section3.DifferentFirstPage = true;

                var section3DefaultHeader = GetSectionDefaultHeader(section3, "section3 default header after enabling page differences");
                var section3DefaultFooter = GetSectionDefaultFooter(section3, "section3 default footer after enabling page differences");
                var section3EvenHeader = GetSectionEvenHeader(section3, "section3 even header after enabling page differences");
                var section3EvenFooter = GetSectionEvenFooter(section3, "section3 even footer after enabling page differences");
                var section3FirstHeader = GetSectionFirstHeader(section3, "section3 first header after enabling page differences");
                var section3FirstFooter = GetSectionFirstFooter(section3, "section3 first footer after enabling page differences");

                Assert.NotNull(section3DefaultHeader);
                Assert.NotNull(section3DefaultFooter);
                Assert.NotNull(section3EvenHeader);
                Assert.NotNull(section3EvenFooter);
                Assert.NotNull(section3FirstHeader);
                Assert.NotNull(section3FirstFooter);

                section3DefaultHeader.AddParagraph().AddText("Section 3 - Header Odd/Default");
                section3DefaultFooter.AddParagraph().AddText("Section 3 - Footer Odd/Default");
                section3EvenHeader.AddParagraph().AddText("Section 3 - Header Even");
                section3EvenFooter.AddParagraph().AddText("Section 3 - Footer Even");

                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 2");
                document.AddPageBreak();
                section3.AddParagraph("Test Last Section - 3");



                GetSectionDefaultHeader(document, 0, "section 0 default header").AddParagraph("Section 0").AddBookmark("BookmarkInSection0Header1");
                var tableHeader = GetSectionDefaultHeader(document, 0, "section 0 default header").AddTable(3, 4);
                tableHeader.Rows[0].Cells[3].Paragraphs[0].Text = "This is sparta";

                GetSectionDefaultHeader(document, 0, "section 0 default header").AddHorizontalLine();
                GetSectionDefaultHeader(document, 0, "section 0 default header").AddHyperLink("Link to website!", new Uri("https://evotec.xyz"));
                GetSectionDefaultHeader(document, 0, "section 0 default header").AddHyperLink("Przemysław Klys Email Me", new Uri("mailto:kontakt@evotec.pl?subject=Test Subject"));
                GetSectionDefaultHeader(document, 0, "section 0 default header").AddField(WordFieldType.Author, WordFieldFormat.FirstCap);

                Assert.True(GetSectionDefaultHeader(document, 0, "section 0 default header").Paragraphs.Count == 8);

                GetSectionDefaultFooter(section2, "section2 default footer").AddParagraph().AddText("Section 0").AddBookmark("BookmarkInSection0Header2");
                var tableFooter = GetSectionDefaultFooter(section2, "section2 default footer").AddTable(2, 3);
                tableFooter.Rows[0].Cells[2].Paragraphs[0].Text = "This is not sparta";
                GetSectionDefaultFooter(section2, "section2 default footer").AddHorizontalLine();
                GetSectionDefaultFooter(section2, "section2 default footer").AddHyperLink("Link to website!", new Uri("https://evotec.pl"));
                GetSectionDefaultFooter(section2, "section2 default footer").AddHyperLink("Przemysław Email Me", new Uri("mailto:contact@evotec.pl?subject=Test Subject"));
                GetSectionDefaultFooter(section2, "section2 default footer").AddField(WordFieldType.Author, WordFieldFormat.FirstCap);

                Assert.True(GetSectionDefaultHeader(document, 0, "section 0 default header").Paragraphs.Count == 8);
                Assert.True(GetSectionDefaultHeader(document, 0, "section 0 default header").ParagraphsHyperLinks.Count == 2);
                Assert.True(GetSectionDefaultHeader(document, 0, "section 0 default header").ParagraphsFields.Count == 1);
                Assert.True(GetSectionDefaultHeader(document, 0, "section 0 default header").Tables.Count == 1);

                Assert.True(GetSectionDefaultFooter(document, 2, "section 2 default footer").Paragraphs.Count == 7);
                Assert.True(GetSectionDefaultFooter(document, 2, "section 2 default footer").ParagraphsHyperLinks.Count == 2);
                Assert.True(GetSectionDefaultFooter(document, 2, "section 2 default footer").ParagraphsFields.Count == 1);
                Assert.True(GetSectionDefaultFooter(document, 2, "section 2 default footer").Tables.Count == 1);

                Assert.True(GetSectionDefaultFooter(section2, "section2 default footer").Paragraphs.Count == 7);
                Assert.True(GetSectionDefaultFooter(section2, "section2 default footer").ParagraphsHyperLinks.Count == 2);
                Assert.True(GetSectionDefaultFooter(section2, "section2 default footer").ParagraphsFields.Count == 1);
                Assert.True(GetSectionDefaultFooter(section2, "section2 default footer").Tables.Count == 1);


                document.Save(false);

                Assert.True(document.Sections[0].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[0].DifferentFirstPage == false);

                Assert.True(document.Sections[1].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[1].DifferentFirstPage == false);

                Assert.True(document.Sections[2].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[2].DifferentFirstPage == false);

                Assert.True(document.Sections[3].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[3].DifferentFirstPage == true);

                Assert.True(GetParagraphAt(document.Sections[3].Paragraphs, 0, "section 3 paragraphs").Text == "Test Last Section - 1");

            }
            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSectionsOddEventFirst.docx"))) {
                Assert.True(document.Sections[0].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[0].DifferentFirstPage == false);

                Assert.True(document.Sections[1].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[1].DifferentFirstPage == false);

                Assert.True(document.Sections[2].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[2].DifferentFirstPage == false);

                Assert.True(document.Sections[3].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[3].DifferentFirstPage == true);

                var evenSection1 = document.Sections[1];
                evenSection1.DifferentOddAndEvenPages = true;
                var section1EvenHeader = GetSectionEvenHeader(evenSection1, "section 1 even header after enabling DifferentOddAndEvenPages");
                var section1EvenFooterReloaded = GetSectionEvenFooter(evenSection1, "section 1 even footer after enabling DifferentOddAndEvenPages");
                Assert.NotNull(section1EvenHeader);
                Assert.NotNull(section1EvenFooterReloaded);

                var firstSection2 = document.Sections[2];
                firstSection2.DifferentFirstPage = true;
                var section2FirstHeader = GetSectionFirstHeader(firstSection2, "section 2 first header after enabling DifferentFirstPage");
                var section2FirstFooter = GetSectionFirstFooter(firstSection2, "section 2 first footer after enabling DifferentFirstPage");
                Assert.NotNull(section2FirstHeader);
                Assert.NotNull(section2FirstFooter);

                Assert.True(document.Sections[0].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[0].DifferentFirstPage == false);

                Assert.True(document.Sections[1].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[1].DifferentFirstPage == false);

                Assert.True(document.Sections[2].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[2].DifferentFirstPage == true);

                Assert.True(document.Sections[3].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[3].DifferentFirstPage == true);
                Assert.True(GetParagraphAt(document.Sections[3].Paragraphs, 0, "section 3 paragraphs").Text == "Test Last Section - 1");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithHeadersAndSectionsOddEventFirst.docx"))) {
                Assert.True(document.Sections[0].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[0].DifferentFirstPage == false);

                Assert.True(document.Sections[1].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[1].DifferentFirstPage == false);

                Assert.True(document.Sections[2].DifferentOddAndEvenPages == false);
                Assert.True(document.Sections[2].DifferentFirstPage == true);

                Assert.True(document.Sections[3].DifferentOddAndEvenPages == true);
                Assert.True(document.Sections[3].DifferentFirstPage == true);

                Assert.True(GetParagraphAt(document.Sections[3].Paragraphs, 0, "section 3 paragraphs").Text == "Test Last Section - 1");

                document.Save();
            }
        }

    }
}
