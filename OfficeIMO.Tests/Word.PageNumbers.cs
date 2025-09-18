using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains page numbering tests.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_AddingPageNumberToParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberParagraph.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var table = footer.AddTable(1, 2);
                table.Rows[0].Cells[0].AddParagraph("Footer");
                var para = table.Rows[0].Cells[1].AddParagraph();
                para.AddPageNumber(includeTotalPages: true);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.NotNull(document.ParagraphsFields);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_PageNumberSettings() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberSettings.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].AddPageNumbering(2, NumberFormatValues.LowerRoman);
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                footer.AddParagraph().AddPageNumber();
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var section = document.Sections[0];
                var pageNumberType = section.PageNumberType;
                Assert.NotNull(pageNumberType);
                var format = pageNumberType.Format;
                Assert.NotNull(format);
                Assert.Equal(NumberFormatValues.LowerRoman, format.Value);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_PageNumberSeparator() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberSeparator.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var para = footer.AddParagraph();
                para.AddPageNumber(includeTotalPages: true, separator: " / ");
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_PageNumberStylesAreValid() {
            foreach (WordPageNumberStyle style in Enum.GetValues(typeof(WordPageNumberStyle))) {
                string filePath = Path.Combine(_directoryWithFiles, $"PageNumberStyle_{style}.docx");
                using (WordDocument document = WordDocument.Create(filePath)) {
                    document.AddHeadersAndFooters();
                    var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                    header.AddPageNumber(style);
                    document.Save(false);
                }

                using (WordDocument document = WordDocument.Load(filePath)) {
                    var errors = document.ValidateDocument();
                    errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue").ToList();
                    Assert.True(errors.Count == 0, $"Style {style} errors: {Word.FormatValidationErrors(errors)}");
                }
            }
        }

        [Fact]
        public void Test_PageNumberWithCustomText() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberCustomText.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var pageNumber = header.AddPageNumber(WordPageNumberStyle.PlainNumber);
                pageNumber.AppendText(" custom");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart = document._wordprocessingDocument.MainDocumentPart;
                Assert.NotNull(mainPart);
                var headerPart = mainPart.HeaderParts.First();
                string text = headerPart.Header!.InnerText;
                Assert.Contains("custom", text);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_PageNumberWithTotalPages() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberTotalPages.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var pageNumber = footer.AddPageNumber(WordPageNumberStyle.PlainNumber);
                pageNumber.AppendText(" of ");
                pageNumber.Paragraph.AddField(WordFieldType.NumPages);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart = document._wordprocessingDocument.MainDocumentPart;
                Assert.NotNull(mainPart);
                var footerPart = mainPart.FooterParts.First();
                string text = footerPart.Footer!.InnerText;
                Assert.Contains(" of ", text);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
      
        [Fact]
        public void Test_PageNumberRestartInNewSection() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionPageNumberReset.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var firstSectionFooter = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                firstSectionFooter.AddParagraph().AddPageNumber();

                document.AddParagraph("Section 1");
                var section = document.AddSection();
                section.AddPageNumbering(1);
                section.AddParagraph("Section 2");
                var secondSectionFooter = RequireSectionFooter(document, 1, HeaderFooterValues.Default);
                secondSectionFooter.AddParagraph().AddPageNumber();

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Sections.Count);
                Assert.NotNull(document.Sections[1].PageNumberType);
                var section1 = document.Sections[1];
                var pageNumberType1 = section1.PageNumberType;
                Assert.NotNull(pageNumberType1);
                var start = pageNumberType1.Start;
                Assert.NotNull(start);
                Assert.Equal(1, start.Value);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_PageNumberRomanFormat() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberRoman.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Sections[0].AddPageNumbering(1, NumberFormatValues.UpperRoman);
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var para = footer.AddParagraph();
                para.AddPageNumber(includeTotalPages: true, format: WordFieldFormat.Roman);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var sectionA = document.Sections[0];
                var pageNumberTypeA = sectionA.PageNumberType;
                Assert.NotNull(pageNumberTypeA);
                var formatA = pageNumberTypeA.Format;
                Assert.NotNull(formatA);
                Assert.Equal(NumberFormatValues.UpperRoman, formatA.Value);
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                Assert.Contains(footer.Fields, f => f.FieldType == WordFieldType.Page && f.FieldFormat.Contains(WordFieldFormat.Roman));
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }

        [Theory]
        [InlineData("0")]
        [InlineData("00")]
        [InlineData("000")]
        [InlineData("0000")]
        [InlineData("#")]
        [InlineData("##")]
        [InlineData("###")]
        [InlineData("#,##0")]
        [InlineData("0.00")]
        [InlineData("##0.##")]
        [InlineData("000#")]
        [InlineData("#000")]
        [InlineData("10-20")]
        [InlineData("Page 0")]
        [InlineData("0-00")]
        public void Test_PageNumberCustomFormat(string format) {
            string filePath = Path.Combine(_directoryWithFiles, $"PageNumberCustomFormat_{Guid.NewGuid()}.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var footer = RequireSectionFooter(document, 0, HeaderFooterValues.Default);
                var pageNumber = footer.AddPageNumber(WordPageNumberStyle.PlainNumber);
                pageNumber.CustomFormat = format;
                Assert.Equal(format, pageNumber.CustomFormat);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                var mainPart2 = document._wordprocessingDocument.MainDocumentPart;
                Assert.NotNull(mainPart2);
                var footerPart = mainPart2.FooterParts.First();
                string xml = footerPart.Footer!.InnerXml;
                Assert.Contains($"\\@ \"{format}\"", xml);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
    }
}
