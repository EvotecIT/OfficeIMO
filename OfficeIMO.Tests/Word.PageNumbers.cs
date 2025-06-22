using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_AddingPageNumberToParagraph() {
            string filePath = Path.Combine(_directoryWithFiles, "PageNumberParagraph.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                var table = document.Footer.Default.AddTable(1, 2);
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
                document.Footer.Default.AddParagraph().AddPageNumber();
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(NumberFormatValues.LowerRoman, document.Sections[0].PageNumberType.Format.Value);
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
                var para = document.Footer.Default.AddParagraph();
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
                    document.Header.Default.AddPageNumber(style);
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
                var pageNumber = document.Header.Default.AddPageNumber(WordPageNumberStyle.PlainNumber);
                pageNumber.Paragraphs.Last().AddText(" custom");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var headerPart = document._wordprocessingDocument.MainDocumentPart.HeaderParts.First();
                string text = headerPart.Header.InnerText;
                Assert.Contains("custom", text);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
    }
}
