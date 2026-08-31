using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const string PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Theory]
    [InlineData("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", false)]
    [InlineData("application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", true)]
    [InlineData("application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", true)]
    [InlineData("application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", true)]
    public void XlsxNativePackage_PreservesSdkBehaviorForUnexpectedContentType(
        string expectedContentType,
        bool sdkRejectsPackage) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            string contentTypes = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformed = contentTypes.Replace(
                expectedContentType,
                "text/plain");
            Assert.NotEqual(contentTypes, malformed);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformed));

            if (sdkRejectsPackage) {
                IOException exception = Assert.Throws<IOException>(() => {
                    using var reader = ExcelDocument.OpenDataReader(path);
                    while (reader.Read()) {
                        _ = reader.GetInt32(0);
                    }
                });
                Assert.IsType<OpenXmlPackageException>(exception.InnerException);
            } else {
                AssertCompactNumericRows(path);
            }
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XlsxNativePackage_RejectsMissingOrDuplicateRootRelationshipId(bool duplicate) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "_rels/.rels";
            XDocument relationships = LoadZipXml(path, entryName);
            XNamespace packageRelationships = PackageRelationshipsNamespace;
            XElement workbookRelationship = relationships.Root!
                .Elements(packageRelationships + "Relationship")
                .Single(element => ((string?)element.Attribute("Type"))?.EndsWith(
                    "/officeDocument",
                    StringComparison.Ordinal) == true);
            if (duplicate) {
                relationships.Root.Add(new XElement(
                    packageRelationships + "Relationship",
                    new XAttribute("Id", (string)workbookRelationship.Attribute("Id")!),
                    new XAttribute("Type", OfficeRelationshipsNamespace + "/metadata/core-properties"),
                    new XAttribute("Target", "docProps/unused.xml")));
            } else {
                workbookRelationship.Attribute("Id")!.Remove();
            }
            ReplaceZipXml(path, entryName, relationships);

            Assert.Throws<IOException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void XlsxNativePackage_RejectsMalformedUnusedContentTypeEntries(bool duplicateOverride) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            XDocument contentTypes = LoadZipXml(path, entryName);
            XNamespace contentTypeNamespace =
                "http://schemas.openxmlformats.org/package/2006/content-types";
            if (duplicateOverride) {
                contentTypes.Root!.Add(
                    new XElement(
                        contentTypeNamespace + "Override",
                        new XAttribute("PartName", "/unused/metadata.xml"),
                        new XAttribute("ContentType", "application/x-officeimo-unused")),
                    new XElement(
                        contentTypeNamespace + "Override",
                        new XAttribute("PartName", "/unused/metadata.xml"),
                        new XAttribute("ContentType", "application/x-officeimo-duplicate")));
            } else {
                contentTypes.Root!.Add(
                    new XElement(
                        contentTypeNamespace + "Default",
                        new XAttribute("Extension", "unused"),
                        new XAttribute("ContentType", "application/x-officeimo-unused")),
                    new XElement(
                        contentTypeNamespace + "Default",
                        new XAttribute("Extension", "unused"),
                        new XAttribute("ContentType", "application/x-officeimo-duplicate")));
            }
            ReplaceZipXml(path, entryName, contentTypes);

            Assert.Throws<ArgumentException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Theory]
    [InlineData("/unused/metadata.xml", "application/x invalid")]
    [InlineData("//unused/metadata.xml", "application/x-officeimo-unused")]
    public void XlsxNativePackage_FallsBackForMalformedUnusedContentTypeDeclaration(
        string partName,
        string contentType) {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            XDocument contentTypes = LoadZipXml(path, entryName);
            XNamespace contentTypeNamespace =
                "http://schemas.openxmlformats.org/package/2006/content-types";
            contentTypes.Root!.Add(new XElement(
                contentTypeNamespace + "Override",
                new XAttribute("PartName", partName),
                new XAttribute("ContentType", contentType)));
            ReplaceZipXml(path, entryName, contentTypes);

            Assert.Throws<XlsxTabularFastPathNotSupportedException>(() =>
                XlsxTabularWorkbook.Open(path, new ExcelReadOptions()));
            Assert.ThrowsAny<ArgumentException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void XlsxNativePackage_UsesNativePathForExplicitSheetIndex() {
        string path = Path.Combine(
            Path.GetTempPath(),
            $"OfficeIMO.Excel.XlsxNativeSheetIndex.{Guid.NewGuid():N}.xlsx");
        try {
            using (var document = ExcelDocument.Create(path)) {
                ExcelSheet first = document.AddWorksheet("First");
                first.CellValue(1, 1, "Value");
                first.CellValue(2, 1, 11);
                ExcelSheet second = document.AddWorksheet("Second");
                second.CellValue(1, 1, "Value");
                second.CellValue(2, 1, 22);
                document.Save();
            }

            using var workbook = XlsxTabularWorkbook.Open(
                path,
                new ExcelReadOptions { SheetIndex = 1 });
            Assert.Equal(new[] { "First", "Second" }, workbook.TableNames);
            using var reader = workbook.OpenTable(
                "Second",
                hasHeaderRow: true,
                CancellationToken.None);
            Assert.True(reader.Read());
            Assert.Equal(22, reader.GetInt32(0));
            Assert.False(reader.Read());
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void XlsxNativePackage_PreservesPercentEncodedWorksheetPartNames() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string originalEntryName = "xl/worksheets/sheet1.xml";
            const string encodedEntryName = "xl/worksheets/sheet%201.xml";
            byte[] worksheet = ReadZipEntry(path, originalEntryName);
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                archive.GetEntry(originalEntryName)!.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(
                    encodedEntryName,
                    CompressionLevel.Optimal);
                using Stream output = replacement.Open();
                output.Write(worksheet, 0, worksheet.Length);
            }

            const string relationshipsEntry = "xl/_rels/workbook.xml.rels";
            XDocument relationships = LoadZipXml(path, relationshipsEntry);
            XNamespace packageRelationships = PackageRelationshipsNamespace;
            XElement worksheetRelationship = relationships.Root!
                .Elements(packageRelationships + "Relationship")
                .Single(element => ((string?)element.Attribute("Type"))?.EndsWith(
                    "/worksheet",
                    StringComparison.Ordinal) == true);
            worksheetRelationship.SetAttributeValue("Target", "worksheets/sheet%201.xml");
            ReplaceZipXml(path, relationshipsEntry, relationships);

            const string contentTypesEntry = "[Content_Types].xml";
            XDocument contentTypes = LoadZipXml(path, contentTypesEntry);
            XNamespace contentTypeNamespace =
                "http://schemas.openxmlformats.org/package/2006/content-types";
            XElement worksheetOverride = contentTypes.Root!
                .Elements(contentTypeNamespace + "Override")
                .Single(element => string.Equals(
                    (string?)element.Attribute("PartName"),
                    "/" + originalEntryName,
                    StringComparison.Ordinal));
            worksheetOverride.SetAttributeValue("PartName", "/" + encodedEntryName);
            ReplaceZipXml(path, contentTypesEntry, contentTypes);

            using (var workbook = XlsxTabularWorkbook.Open(path, new ExcelReadOptions())) {
                Assert.Equal(new[] { "Data" }, workbook.TableNames);
                using var reader = workbook.OpenTable(
                    "Data",
                    hasHeaderRow: true,
                    CancellationToken.None);
                Assert.True(reader.Read());
                Assert.Equal(42, reader.GetInt32(0));
                Assert.True(reader.Read());
                Assert.Equal(43, reader.GetInt32(0));
                Assert.False(reader.Read());
            }

            Assert.Equal(new[] { "Data" }, ExcelDocument.GetSheetNames(path));
        } finally {
            File.Delete(path);
        }
    }

#if !NETFRAMEWORK
    [Fact]
    public void XlsxNativePackage_FallsBackForWorksheetLargerThanNativeBuffer() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/worksheets/sheet1.xml";
            byte[] worksheet = ReadZipEntry(path, entryName);
            byte[] closingTag = "</worksheet>"u8.ToArray();
            int closingTagOffset = worksheet.AsSpan().LastIndexOf(closingTag);
            Assert.True(closingTagOffset >= 0);

            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                ZipArchiveEntry original = archive.GetEntry(entryName)!;
                original.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using Stream destination = replacement.Open();
                destination.Write(worksheet, 0, closingTagOffset);
                byte[] padding = Enumerable.Repeat((byte)' ', 64 * 1024).ToArray();
                for (int index = 0; index <= 1024; index++) {
                    destination.Write(padding, 0, padding.Length);
                }
                destination.Write(worksheet, closingTagOffset, worksheet.Length - closingTagOffset);
            }

            AssertCompactNumericRows(path);
        } finally {
            File.Delete(path);
        }
    }
#endif

    [Fact]
    public void XlsxNativePackage_RejectsUnknownWorkbookRelationshipTargetMode() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "xl/_rels/workbook.xml.rels";
            XDocument relationships = LoadZipXml(path, entryName);
            XNamespace packageRelationships = PackageRelationshipsNamespace;
            XElement worksheetRelationship = relationships.Root!
                .Elements(packageRelationships + "Relationship")
                .Single(element => ((string?)element.Attribute("Type"))?.EndsWith(
                    "/worksheet",
                    StringComparison.Ordinal) == true);
            worksheetRelationship.SetAttributeValue("TargetMode", "Bogus");
            ReplaceZipXml(path, entryName, relationships);

            IOException exception = Assert.Throws<IOException>(() => ExcelDocument.OpenDataReader(path));
            var xmlException = Assert.IsType<System.Xml.XmlException>(exception.InnerException);
            Assert.IsType<ArgumentException>(xmlException.InnerException);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void XlsxNativePackage_RejectsMissingNonWorksheetSheetPart() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string relationshipId = "rIdMissingChart";
            const string relationshipsEntry = "xl/_rels/workbook.xml.rels";
            XDocument relationships = LoadZipXml(path, relationshipsEntry);
            XNamespace packageRelationships = PackageRelationshipsNamespace;
            relationships.Root!.Add(new XElement(
                packageRelationships + "Relationship",
                new XAttribute("Id", relationshipId),
                new XAttribute("Type", OfficeRelationshipsNamespace + "/chartsheet"),
                new XAttribute("Target", "chartsheets/missing.xml")));
            ReplaceZipXml(path, relationshipsEntry, relationships);

            const string workbookEntry = "xl/workbook.xml";
            XDocument workbook = LoadZipXml(path, workbookEntry);
            XNamespace spreadsheet = workbook.Root!.Name.Namespace;
            XNamespace officeRelationships = OfficeRelationshipsNamespace;
            workbook.Root.Element(spreadsheet + "sheets")!.Add(new XElement(
                spreadsheet + "sheet",
                new XAttribute("name", "MissingChart"),
                new XAttribute("sheetId", "2"),
                new XAttribute(officeRelationships + "id", relationshipId)));
            ReplaceZipXml(path, workbookEntry, workbook);

            Assert.Throws<InvalidOperationException>(() => ExcelDocument.OpenDataReader(path));
        } finally {
            File.Delete(path);
        }
    }

#if !NETFRAMEWORK
    [Fact]
    public void XlsxNativePackage_IgnoresUnusedZipEntryUnsupportedByNativeIndex() {
        string path = CreateCompactFastPathWorkbook();
        try {
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                archive.CreateEntry(" ");
            }

            AssertCompactNumericRows(path);
        } finally {
            File.Delete(path);
        }
    }
#endif

    private static XDocument LoadZipXml(string path, string entryName) =>
        XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(path, entryName)));

    private static void ReplaceZipXml(string path, string entryName, XDocument document) =>
        ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting)));
}
