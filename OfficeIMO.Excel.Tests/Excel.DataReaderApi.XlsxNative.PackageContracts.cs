using DocumentFormat.OpenXml.Packaging;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Excel.Tests;

public partial class Excel {
    private const string PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string OfficeRelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    [Fact]
    public void XlsxNativePackage_RejectsWorksheetWithUnexpectedContentType() {
        string path = CreateCompactFastPathWorkbook();
        try {
            const string entryName = "[Content_Types].xml";
            string contentTypes = Encoding.UTF8.GetString(ReadZipEntry(path, entryName));
            string malformed = contentTypes.Replace(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
                "text/plain");
            Assert.NotEqual(contentTypes, malformed);
            ReplaceZipEntry(path, entryName, Encoding.UTF8.GetBytes(malformed));

            IOException exception = Assert.Throws<IOException>(() => ExcelDocument.OpenDataReader(path));
            Assert.IsType<OpenXmlPackageException>(exception.InnerException);
        } finally {
            File.Delete(path);
        }
    }

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
