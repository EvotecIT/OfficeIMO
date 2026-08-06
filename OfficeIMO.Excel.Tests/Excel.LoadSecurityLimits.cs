using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LoadAppliesOpenXmlCharacterLimitBeforeContentTypeNormalization() {
            byte[] package;
            using (var output = new MemoryStream()) {
                using (ExcelDocument document = ExcelDocument.Create()) {
                    document.AddWorksheet("Data").CellValue(1, 1, "safe");
                    document.Save(output);
                }
                package = output.ToArray();
            }

            using var input = new MemoryStream(package, writable: false);
            var options = new ExcelLoadOptions {
                AccessMode = DocumentAccessMode.ReadOnly,
                OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 128 }
            };

            IOException exception = Assert.Throws<IOException>(() => ExcelDocument.Load(input, options));

            InvalidDataException inner = Assert.IsType<InvalidDataException>(exception.InnerException);
            Assert.Contains("content-types part", inner.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("128-character limit", inner.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void OpenFileBackedAppliesOpenXmlCharacterLimitBeforeContentTypeNormalization() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            try {
                using (ExcelDocument document = ExcelDocument.Create()) {
                    document.AddWorksheet("Data").CellValue(1, 1, "safe");
                    document.Save(path);
                }
                var options = new ExcelLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    OpenSettings = new OfficeOpenXmlLoadSettings { MaxCharactersInPart = 128 }
                };

                InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
                    ExcelDocument.OpenFileBacked(path, options));

                Assert.Contains("content-types part", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("128-character limit", exception.Message, StringComparison.Ordinal);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
