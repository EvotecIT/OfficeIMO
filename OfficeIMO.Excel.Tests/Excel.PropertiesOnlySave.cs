using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Threading.Tasks;
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
                sheet.Cell(1, 1).SetValue("Name");
                sheet.Cell(2, 1).SetValue(42);
                document.Properties.Subject = "Kept";
                document.Save(source);
            }
            using (var package = SpreadsheetDocument.Open(source, isEditable: true)) {
                package.PackageProperties.ContentStatus = "Draft";
            }
            byte[] original = source.ToArray();

            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(original))) {
                document.Properties.Title = "New title";
                document.Properties.Category = "New category";
                document.Save(output);
                Assert.Equal("New title", document.Properties.Title);
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
                document.AddSheet("Data").Cell(1, 1).SetValue("Before");
                document.Save(source);
            }

            using var output = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(source.ToArray()))) {
                document.Properties.Title = "Title";
                document.Sheets[0].Cell(1, 1).SetValue("After");
                document.Save(output);
            }

            output.Position = 0;
            using var reopened = ExcelDocument.Load(output);
            Assert.Equal("Title", reopened.Properties.Title);
            Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string text));
            Assert.Equal("After", text);
        }

        [Fact]
        public void Saving_a_loaded_workbook_twice_keeps_cell_edits_and_added_sheets() {
            using var source = new MemoryStream();
            using (var document = ExcelDocument.Create(source)) {
                document.AddSheet("Edited").Cell(1, 1).SetValue("Before");
                document.AddSheet("Untouched").Cell(1, 1).SetValue("Same");
                document.Save(source);
            }
            byte[] original = source.ToArray();

            using var cellEdit = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(original))) {
                document.Sheets[0].Cell(1, 1).SetValue("After");
                document.Save(cellEdit);
            }

            using var added = new MemoryStream();
            using (var document = ExcelDocument.Load(new MemoryStream(cellEdit.ToArray()))) {
                document.AddSheet("Added").Cell(1, 1).SetValue("New");
                document.Sheets[1].Cell(2, 1).SetValue("Also");
                document.Save(added);
            }

            added.Position = 0;
            using (var package = SpreadsheetDocument.Open(added, isEditable: false)) {
                Assert.Empty(new OpenXmlValidator().Validate(package));
            }
            added.Position = 0;
            using var reopened = ExcelDocument.Load(added);
            Assert.Equal(new[] { "Edited", "Untouched", "Added" }, reopened.Sheets.Select(sheet => sheet.Name));
            Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string edited));
            Assert.Equal("After", edited);
            Assert.True(reopened.Sheets[1].TryGetCellText(2, 1, out string also));
            Assert.Equal("Also", also);
            Assert.True(reopened.Sheets[2].TryGetCellText(1, 1, out string created));
            Assert.Equal("New", created);
        }

        [Theory]
        [InlineData(false, false)]
        [InlineData(true, false)]
        [InlineData(false, true)]
        [InlineData(true, true)]
        public async Task Saving_a_loaded_workbook_leaves_unchanged_entries_as_the_source_had_them(bool toFile, bool asynchronous) {
            using var authored = new MemoryStream();
            using (var document = ExcelDocument.Create(authored)) {
                document.AddSheet("Edited").Cell(1, 1).SetValue("Before");
                document.Save(authored);
            }
            string relationships = "xl/_rels/workbook.xml.rels";
            byte[] expected = ReadEntry(authored.ToArray(), relationships);

            string path = Path.Combine(_directoryWithFiles, $"SplicedRoute-{toFile}-{asynchronous}.xlsx");
            File.WriteAllBytes(path, authored.ToArray());
            using var target = new MemoryStream();
            using (var document = ExcelDocument.Load(path)) {
                document.Sheets[0].Cell(1, 1).SetValue("After");
                if (toFile && asynchronous) await document.SaveAsync(path);
                else if (toFile) document.Save(path);
                else if (asynchronous) await document.SaveAsync(target);
                else document.Save(target);
            }

            byte[] saved = toFile ? File.ReadAllBytes(path) : target.ToArray();
            Assert.Equal(expected, ReadEntry(saved, relationships));
            using var reopened = ExcelDocument.Load(new MemoryStream(saved));
            Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string text));
            Assert.Equal("After", text);
            File.Delete(path);
        }

        private static byte[] ReadEntry(byte[] package, string name) {
            using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
            using Stream entry = archive.GetEntry(name)!.Open();
            using var content = new MemoryStream();
            entry.CopyTo(content);
            return content.ToArray();
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
