using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Changing_only_properties_keeps_every_other_package_entry_and_all_other_core_values() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create(source)) {
                ExcelSheet sheet = document.AddSheet("Data");
                sheet.CellValue(1, 1, "Name");
                sheet.CellValue(2, 1, 42);
                document.BuiltinDocumentProperties.Subject = "Kept";
                document.Save(source);
            }
            using (var package = SpreadsheetDocument.Open(source, isEditable: true)) {
                package.PackageProperties.ContentStatus = "Draft";
            }
            byte[] original = source.ToArray();

            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(original))) {
                document.BuiltinDocumentProperties.Title = "New title";
                document.BuiltinDocumentProperties.Category = "New category";
                document.Save(output);
                Assert.Equal("New title", document.BuiltinDocumentProperties.Title);
            }

            output.Position = 0;
            using (var package = SpreadsheetDocument.Open(output, isEditable: false)) {
                Assert.Empty(new OpenXmlValidator().Validate(package));
                Assert.Equal("New title", package.PackageProperties.Title);
                Assert.Equal("New category", package.PackageProperties.Category);
                Assert.Equal("Kept", package.PackageProperties.Subject);
                Assert.Equal("Draft", package.PackageProperties.ContentStatus);
            }

            using var before = new ZipArchive(new MemoryStream(original));
            using var after = new ZipArchive(new MemoryStream(output.ToArray()));
            Assert.Equal(
                before.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal),
                after.Entries.Select(entry => entry.FullName).OrderBy(name => name, StringComparer.Ordinal));
            foreach (ZipArchiveEntry entry in before.Entries.Where(entry => !IsCoreProperties(entry))) {
                Assert.Equal(Hash(entry), Hash(after.GetEntry(entry.FullName)!));
            }
        }

        [Fact]
        public void A_later_cell_edit_is_not_lost_when_properties_changed_first() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create(source)) {
                document.AddSheet("Data").CellValue(1, 1, "Before");
                document.Save(source);
            }

            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(source.ToArray()))) {
                document.BuiltinDocumentProperties.Title = "Title";
                document.Sheets[0].CellValue(1, 1, "After");
                document.Save(output);
            }

            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            Assert.Equal("Title", reopened.BuiltinDocumentProperties.Title);
            Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string text));
            Assert.Equal("After", text);
        }

        private static bool IsCoreProperties(ZipArchiveEntry entry) =>
            entry.FullName.EndsWith("core.xml", StringComparison.Ordinal) || entry.FullName.EndsWith(".psmdcp", StringComparison.Ordinal);

        private static string Hash(ZipArchiveEntry entry) {
            using Stream stream = entry.Open();
            using var sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(stream));
        }
    }
}
