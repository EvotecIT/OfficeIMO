using System;
using System.IO;
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
    }
}
