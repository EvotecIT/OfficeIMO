using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
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

            using MemoryStream stream = new System.IO.MemoryStream(document.ToBytes());
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
        }

        [Fact]
        public void Test_TableLookConditionalFormatting_UsesNativeDefaultsAndCollectiveExpandedFlags() {
            using WordDocument document = WordDocument.Create();
            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableLook = null;
            Assert.True(table.ConditionalFormattingFirstRow);
            Assert.False(table.ConditionalFormattingLastRow);
            Assert.True(table.ConditionalFormattingFirstColumn);
            Assert.False(table.ConditionalFormattingLastColumn);
            Assert.False(table.ConditionalFormattingNoHorizontalBand);
            Assert.True(table.ConditionalFormattingNoVerticalBand);
            Assert.Null(table._tableProperties.TableLook);
            table.ConditionalFormattingLastRow = true;
            Assert.Equal("04E0", table._tableProperties.TableLook!.Val!.Value);

            // Sparse expanded attributes replace packed val and leave all omitted flags on.
            table._tableProperties.TableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook {
                Val = "0000", FirstRow = false, FirstColumn = true, NoVerticalBand = true
            };
            Assert.False(table.ConditionalFormattingFirstRow);
            Assert.True(table.ConditionalFormattingLastRow);
            Assert.True(table.ConditionalFormattingFirstColumn);
            Assert.True(table.ConditionalFormattingLastColumn);
            Assert.True(table.ConditionalFormattingNoHorizontalBand);
            Assert.True(table.ConditionalFormattingNoVerticalBand);
            ValidateSaved(FileFormatVersions.Office2010);

            table.ConditionalFormattingLastRow = true;
            Assert.False(table.ConditionalFormattingFirstRow);
            Assert.True(table.ConditionalFormattingLastRow);
            using (MemoryStream stream = new MemoryStream(document.ToBytes()))
            using (WordprocessingDocument package = WordprocessingDocument.Open(stream, false)) {
                var errors = new OpenXmlValidator().Validate(package).ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
                var look = package.MainDocumentPart!.Document.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.TableLook>().Single();
                Assert.Equal("07C0", look.Val!.Value);
                Assert.Null(look.FirstRow);
                Assert.Null(look.LastRow);
            }

            table._tableProperties.TableLook = new DocumentFormat.OpenXml.Wordprocessing.TableLook {
                Val = "07E0", FirstRow = false, LastRow = false, FirstColumn = false,
                LastColumn = false, NoHorizontalBand = false, NoVerticalBand = false
            };
            Assert.False(table.ConditionalFormattingFirstRow);
            Assert.False(table.ConditionalFormattingLastRow);
            Assert.False(table.ConditionalFormattingFirstColumn);
            Assert.False(table.ConditionalFormattingLastColumn);
            Assert.False(table.ConditionalFormattingNoHorizontalBand);
            Assert.False(table.ConditionalFormattingNoVerticalBand);
            ValidateSaved(FileFormatVersions.Office2010);

            void ValidateSaved(FileFormatVersions version) {
                using var stream = new MemoryStream(document.ToBytes());
                using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
                var errors = new OpenXmlValidator(version).Validate(package).ToList();
                Assert.True(errors.Count == 0, Word.FormatValidationErrors(errors));
            }
        }
    }
}
