using System;
using System.IO;
using System.IO.Compression;
using System.Threading.Tasks;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_LoadMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missing.xlsx");
            var ex = Assert.Throws<FileNotFoundException>(() => ExcelDocument.Load(filePath));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public void Test_LoadNullPath_ThrowsArgumentNullException() {
            var ex = Assert.Throws<ArgumentNullException>(() => ExcelDocument.Load((string)null!));
            Assert.Equal("filePath", ex.ParamName);
        }

        [Fact]
        public async Task Test_LoadAsyncMissingFile_ThrowsWithPath() {
            string filePath = Path.Combine(_directoryWithFiles, "missingAsync.xlsx");
            var ex = await Assert.ThrowsAsync<FileNotFoundException>(() => ExcelDocument.LoadAsync(filePath));
            Assert.Equal($"File '{filePath}' doesn't exist.", ex.Message);
        }

        [Fact]
        public async Task Test_LoadAsyncNullPath_ThrowsArgumentNullException() {
            var ex = await Assert.ThrowsAsync<ArgumentNullException>(() => ExcelDocument.LoadAsync((string)null!));
            Assert.Equal("filePath", ex.ParamName);
        }

        [Fact]
        public void Test_LoadInvalidAppPropsContentType_ThrowsHelpfulIOException()
        {
            string sourcePath = Path.Combine(_directoryDocuments, "BasicExcel.xlsx");
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            File.Copy(sourcePath, filePath, overwrite: true);

            using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Update, leaveOpen: false))
            {
                Assert.NotNull(archive.GetEntry("docProps/app.xml"));
                var contentTypes = archive.GetEntry("[Content_Types].xml") ?? throw new InvalidOperationException("Missing content types.");
                contentTypes.Delete();
                var replacement = archive.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
                using var replacementStream = replacement.Open();
                using var writer = new StreamWriter(replacementStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                writer.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
                writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n");
                writer.Write("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n");
                writer.Write("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n");
                writer.Write("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n");
                writer.Write("  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n");
                writer.Write("  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/xml\">");
                // Intentionally omit the closing tag to mimic a corrupted declaration that still references /docProps/app.xml
            }

            var exception = Assert.Throws<IOException>(() => ExcelDocument.Load(filePath));
            Assert.Contains("invalid content type for '/docProps/app.xml'", exception.Message);
            Assert.IsType<XmlException>(exception.InnerException);
        }

        [Fact]
        public void Test_LoadNormalizedPathWithAutoSave_PersistsChangesOnDispose() {
            string sourcePath = Path.Combine(_directoryDocuments, "BasicExcel.xlsx");
            string filePath = Path.Combine(_directoryWithFiles, "LoadNormalizedAutoSave.xlsx");
            File.Copy(sourcePath, filePath, overwrite: true);

            RewriteContentTypes(filePath, root => {
                XNamespace ns = root.Name.Namespace;
                var appOverride = root.Elements(ns + "Override")
                    .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), "/docProps/app.xml", StringComparison.OrdinalIgnoreCase))
                    ?? throw new InvalidOperationException("Missing /docProps/app.xml override.");
                appOverride.SetAttributeValue("ContentType", "application/xml");
            });

            using (var document = ExcelDocument.Load(filePath, readOnly: false, autoSave: true)) {
                document.Sheets[0].CellValue(1, 1, "Normalized");
            }

            byte[] savedBytes = File.ReadAllBytes(filePath);
            using var memory = new MemoryStream(savedBytes);
            using (var reloaded = ExcelDocument.Load(memory)) {
                Assert.True(reloaded.Sheets[0].TryGetCellText(1, 1, out var value));
                Assert.Equal("Normalized", value);
            }
        }

        private static void RewriteContentTypes(string filePath, Action<XElement> mutateRoot) {
            using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            using var archive = new ZipArchive(fileStream, ZipArchiveMode.Update, leaveOpen: false);

            var contentTypes = archive.GetEntry("[Content_Types].xml") ?? throw new InvalidOperationException("Missing content types.");

            string xml;
            using (var entryStream = contentTypes.Open())
            using (var reader = new StreamReader(entryStream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))) {
                xml = reader.ReadToEnd();
            }

            var document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            var root = document.Root ?? throw new InvalidOperationException("Missing content types root.");
            mutateRoot(root);

            contentTypes.Delete();
            var replacement = archive.CreateEntry("[Content_Types].xml", CompressionLevel.NoCompression);
            using var replacementStream = replacement.Open();
            var settings = new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false,
                NewLineHandling = NewLineHandling.None
            };
            using var writer = XmlWriter.Create(replacementStream, settings);
            document.Save(writer);
            writer.Flush();
        }

    }
}
