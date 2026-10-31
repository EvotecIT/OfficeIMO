using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelContentTypeDefaultsTests {
    private const string CustomPropertiesContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
    private const string WorksheetContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Load_preserves_parts_inheriting_a_specialized_xml_default(bool worksheetDefault) {
        byte[] fixture = CreateWorkbookWithInheritedXmlContentType(worksheetDefault);
        AssertSdkWorkbook(fixture, 42, "Original metadata");

        using var input = new MemoryStream();
        input.Write(fixture, 0, fixture.Length);
        input.Position = 0;
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Load(input, new ExcelLoadOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit
        })) {
            ExcelSheet sheet = document["Data"];
            ExcelCustomProperty property = document.CustomDocumentProperties["Review"];
            Assert.Equal(42, sheet.CellAt(1, 1).GetValue<int>());
            Assert.Equal("Original metadata", property.Value);

            sheet.CellValue(1, 1, 84);
            property.Value = "Updated metadata";
            document.BuiltinDocumentProperties.Title = "Updated title";
            document.Save(output);
        }

        Assert.True(input.CanRead);
        Assert.Equal(fixture, input.ToArray());
        byte[] saved = output.ToArray();
        AssertSdkWorkbook(saved, 84, "Updated metadata", "Updated title");

        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        Assert.Equal(84, reopened["Data"].CellAt(1, 1).GetValue<int>());
        Assert.Equal("Updated metadata", reopened.CustomDocumentProperties["Review"].Value);
        Assert.Equal("Updated title", reopened.BuiltinDocumentProperties.Title);
    }

    private static byte[] CreateWorkbookWithInheritedXmlContentType(bool worksheetDefault) {
        using var package = new MemoryStream();
        using (var document = ExcelDocument.Create()) {
            document.AddSheet("Data").CellValue(1, 1, 42);
            document.SetCustomDocumentProperty("Review", "Original metadata");
            document.Save(package);
        }

        string inheritedPartName;
        string inheritedContentType;
        using (var original = new MemoryStream(package.ToArray(), writable: false))
        using (var workbook = SpreadsheetDocument.Open(original, false)) {
            OpenXmlPart part = worksheetDefault
                ? workbook.WorkbookPart!.WorksheetParts.Single()
                : workbook.CustomFilePropertiesPart!;
            inheritedPartName = part.Uri.OriginalString;
            inheritedContentType = part.ContentType;
        }

        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("[Content_Types].xml")!;
            XDocument types;
            using (Stream source = entry.Open()) types = XDocument.Load(source);
            XElement root = types.Root!;
            XNamespace ns = root.Name.Namespace;
            var defaults = root.Elements(ns + "Default").ToDictionary(
                item => item.Attribute("Extension")!.Value,
                item => item.Attribute("ContentType")!.Value,
                StringComparer.OrdinalIgnoreCase);
            var overrides = root.Elements(ns + "Override").ToDictionary(
                item => item.Attribute("PartName")!.Value,
                item => item.Attribute("ContentType")!.Value,
                StringComparer.OrdinalIgnoreCase);

            // Change only how MIME types are declared; every existing XML part keeps its effective type.
            foreach (ZipArchiveEntry part in archive.Entries.Where(item =>
                item.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && item.FullName != "[Content_Types].xml")) {
                string partName = "/" + part.FullName;
                if (string.Equals(partName, inheritedPartName, StringComparison.OrdinalIgnoreCase)
                    || overrides.ContainsKey(partName)) continue;
                root.Add(new XElement(ns + "Override",
                    new XAttribute("PartName", partName),
                    new XAttribute("ContentType", defaults["xml"])));
            }

            root.Elements(ns + "Override").Where(item => string.Equals(
                (string?)item.Attribute("PartName"), inheritedPartName, StringComparison.OrdinalIgnoreCase)).Remove();
            root.Elements(ns + "Default").Where(item => string.Equals(
                (string?)item.Attribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase)).Remove();
            root.AddFirst(new XElement(ns + "Default",
                new XAttribute("Extension", "xml"),
                new XAttribute("ContentType", inheritedContentType)));

            entry.Delete();
            using Stream replacement = archive.CreateEntry("[Content_Types].xml").Open();
            types.Save(replacement);
        }

        return package.ToArray();
    }

    private static void AssertSdkWorkbook(byte[] bytes, int value, string review, string? title = null) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var workbook = SpreadsheetDocument.Open(stream, false);
        WorksheetPart sheet = workbook.WorkbookPart!.WorksheetParts.Single();
        Assert.Equal(WorksheetContentType, sheet.ContentType);
        Assert.Equal(value.ToString(System.Globalization.CultureInfo.InvariantCulture), sheet.Worksheet
            .Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1").CellValue!.Text);
        CustomFilePropertiesPart properties = workbook.CustomFilePropertiesPart!;
        Assert.Equal(CustomPropertiesContentType, properties.ContentType);
        Assert.Equal(review, properties.Properties!.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>()
            .Single(property => property.Name?.Value == "Review").VTLPWSTR!.Text);
        if (title != null) Assert.Equal(title, workbook.PackageProperties.Title);
    }
}
