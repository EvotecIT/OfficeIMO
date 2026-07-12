using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        // Regression test for https://github.com/EvotecIT/OfficeIMO/issues/85
        // Uses the DocumentValidationErrors property to confirm no duplicate table styles
        [Fact]
        public void Test_TableStyles_NoDuplicateValidationErrors() {
            string filePath = Path.Combine(_directoryWithFiles, "TableStylesValidation.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                foreach (WordTableStyle style in Enum.GetValues(typeof(WordTableStyle))) {
                    document.AddTable(1, 1, style);
                }
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.True(document.DocumentValidationErrors.Count == 0,
                    Word.FormatValidationErrors(document.DocumentValidationErrors));
            }
        }

        [Fact]
        public void Test_TableLookConditionalFormatting_SavesAsValidOpenXml() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);
            table.ConditionalFormattingFirstRow = true;
            table.ConditionalFormattingLastRow = true;
            table.ConditionalFormattingFirstColumn = false;
            table.ConditionalFormattingLastColumn = true;
            table.ConditionalFormattingNoHorizontalBand = true;
            table.ConditionalFormattingNoVerticalBand = false;

            Assert.True(table.ConditionalFormattingFirstRow);
            Assert.True(table.ConditionalFormattingLastRow);
            Assert.False(table.ConditionalFormattingFirstColumn);
            Assert.True(table.ConditionalFormattingLastColumn);
            Assert.True(table.ConditionalFormattingNoHorizontalBand);
            Assert.False(table.ConditionalFormattingNoVerticalBand);

            using MemoryStream stream = document.ToStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void Test_TableLookConditionalFormatting_HonorsExpandedFalseOverrides() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook {
                Val = "04A0",
                FirstRow = false,
                FirstColumn = true,
                NoVerticalBand = true
            };

            Assert.False(table.ConditionalFormattingFirstRow);
            Assert.True(table.ConditionalFormattingFirstColumn);
            Assert.True(table.ConditionalFormattingNoVerticalBand);

            table.ConditionalFormattingLastRow = true;

            Assert.False(table.ConditionalFormattingFirstRow);
            Assert.True(table.ConditionalFormattingLastRow);

            using MemoryStream stream = document.ToStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var tableLook = package.MainDocumentPart!.Document.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().Single();
            Assert.Equal("04C0", tableLook.Val!.Value);
        }
    }
}
