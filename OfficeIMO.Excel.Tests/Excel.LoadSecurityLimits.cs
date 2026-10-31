using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Utilities;
using System.IO.Compression;
using System.Text;
using System.Xml;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void RejectedRootRelationshipsDoNotEnterPackageUpdateMode() {
            using var package = new MemoryStream();
            using (ExcelDocument document = ExcelDocument.Create()) {
                document.AddSheet("Data").CellValue(1, 1, "safe");
                document.Save(package);
            }
            using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                archive.GetEntry("_rels/.rels")!.Delete();
                using var output = new StreamWriter(archive.CreateEntry("_rels/.rels").Open(), Encoding.UTF8);
                output.Write("<!DOCTYPE Relationships [<!ENTITY x 'x'>]><Relationships>");
                output.Write(new string(' ', 64 * 1024));
                output.Write("</Relationships>");
            }

            package.Position = 0;
            InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
                ExcelPackageUtilities.NormalizeContentTypes(package, leaveOpen: true, maxCharactersInPart: 8192));
            Assert.IsType<XmlException>(error.InnerException);
        }

        [Fact]
        public void LoadAppliesOpenXmlCharacterLimitBeforeContentTypeNormalization() {
            byte[] package;
            using (var output = new MemoryStream()) {
                using (ExcelDocument document = ExcelDocument.Create()) {
                    document.AddSheet("Data").CellValue(1, 1, "safe");
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
                    document.AddSheet("Data").CellValue(1, 1, "safe");
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
