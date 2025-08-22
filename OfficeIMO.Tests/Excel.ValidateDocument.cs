using System;
using System.IO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ValidateDocument() {
            string filePath = Path.Combine(_directoryWithFiles, "ValidateDocument.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Test");
                document.Save();
            }

            using (var document = ExcelDocument.Load(filePath)) {
                var errors = document.ValidateDocument();
                Assert.True(errors.Count == 0, Excel.FormatValidationErrors(errors));
                Assert.True(document.DocumentIsValid);
            }
        }
    }
}

