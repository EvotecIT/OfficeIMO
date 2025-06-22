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
        public void Test_PageNumberRestartInNewSection() {
            string filePath = Path.Combine(_directoryWithFiles, "SectionPageNumberReset.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeadersAndFooters();
                document.Footer.Default.AddParagraph().AddPageNumber();

                document.AddParagraph("Section 1");
                var section = document.AddSection();
                section.AddPageNumbering(1);
                section.AddParagraph("Section 2");
                document.Footer.Default.AddParagraph().AddPageNumber();

                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(2, document.Sections.Count);
                Assert.NotNull(document.Sections[1].PageNumberType);
                Assert.Equal(1, document.Sections[1].PageNumberType.Start.Value);
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
                var para = document.Footer.Default.AddParagraph();
                para.AddPageNumber(includeTotalPages: true, format: WordFieldFormat.Roman);
                document.Save(false);
            }
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(NumberFormatValues.UpperRoman, document.Sections[0].PageNumberType.Format.Value);
                Assert.Contains(document.Sections[0].Footer.Default.Fields, f => f.FieldType == WordFieldType.Page && f.FieldFormat == WordFieldFormat.Roman);
                var errors = document.ValidateDocument();
                errors = errors.Where(e => e.Id != "Sem_UniqueAttributeValue" && e.Id != "Sch_UnexpectedElementContentExpectingComplex").ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
    }
}
