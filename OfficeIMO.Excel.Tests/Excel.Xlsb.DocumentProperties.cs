using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    private static readonly DateTime XlsbPropertiesCreated = new DateTime(2024, 3, 5, 6, 7, 8, DateTimeKind.Utc);

    [Theory]
    [InlineData("new")]
    [InlineData("xlsx")]
    [InlineData("xlsx-edited")]
    [InlineData("xls")]
    public void Xlsb_SaveWritesDocumentPropertiesFromEverySourceAndLoadReadsThemBack(string source) {
        string xlsbPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        string? sourcePath = source == "new" ? null : Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + (source == "xls" ? ".xls" : ".xlsx"));
        using (ExcelDocument authored = ExcelDocument.Create()) {
            authored.AddSheet("Data").CellValue(1, 1, "Alpha");
            authored.Sheets[0].CellValue(2, 1, 42d);
            SetXlsbTestProperties(authored);
            if (sourcePath == null) {
                authored.Save(xlsbPath);
            } else {
                authored.Save(sourcePath);
            }
        }

        if (sourcePath != null) {
            using ExcelDocument loaded = ExcelDocument.Load(sourcePath);
            if (source == "xlsx-edited") loaded.Properties.Keywords = "edited";
            loaded.Save(xlsbPath);
        }

        using ExcelDocument reloaded = ExcelDocument.Load(xlsbPath);
        Assert.Equal(ExcelFileFormat.Xlsb, reloaded.SourceFormat);
        Assert.Equal("Alpha", reloaded.Sheets[0].Cell(1, 1).GetValue<string>());
        Assert.Equal(42d, reloaded.Sheets[0].Cell(2, 1).GetValue<double>());
        Assert.Equal("Quarterly figures", reloaded.Properties.Title);
        Assert.Equal("Finance", reloaded.Properties.Subject);
        Assert.Equal("Report Author", reloaded.Properties.Creator);
        Assert.Equal(source == "xlsx-edited" ? "edited" : "budget", reloaded.Properties.Keywords);
        Assert.Equal("Figures for the quarter", reloaded.Properties.Description);
        Assert.Equal("Report Editor", reloaded.Properties.LastModifiedBy);
        Assert.Equal("Reports", reloaded.Properties.Category);
        Assert.Equal(XlsbPropertiesCreated, reloaded.Properties.Created!.Value.ToUniversalTime());
        Assert.NotNull(reloaded.Properties.Modified);
        Assert.Equal("Example Company", reloaded.ApplicationProperties.Company);
        Assert.Equal("Example Manager", reloaded.ApplicationProperties.Manager);
        Assert.Equal("North", reloaded.CustomProperties["Region"].Text);
        Assert.Equal(7, reloaded.CustomProperties["Revision Count"].NumberInteger);
    }

    [Fact]
    public void Xlsb_ResaveKeepsLoadedPropertiesAndWritesPropertyEditsOnce() {
        string sourcePath = Path.Combine(AppContext.BaseDirectory, "Documents", "XlsbCorpus", "excel-generated", "basic-values-formula.xlsb");
        string cellEditPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        string propertyEditPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        var excelCreated = new DateTime(2026, 7, 16, 7, 48, 15, DateTimeKind.Utc);
        string? excelCreator;

        using (ExcelDocument document = ExcelDocument.Load(sourcePath)) {
            excelCreator = document.Properties.Creator;
            Assert.False(string.IsNullOrEmpty(excelCreator));
            Assert.Equal(excelCreated, document.Properties.Created!.Value.ToUniversalTime());
            Assert.Equal("Microsoft Macintosh Excel", document.ApplicationProperties.ApplicationName);
            document.Sheets[0].CellValue(2, 1, "Edited");
            document.Save(cellEditPath);
        }

        using (ExcelDocument document = ExcelDocument.Load(cellEditPath)) {
            Assert.Equal("Edited", document.Sheets[0].Cell(2, 1).GetValue<string>());
            Assert.Equal(excelCreator, document.Properties.Creator);
            Assert.Equal(excelCreated, document.Properties.Created!.Value.ToUniversalTime());
            document.Properties.Title = "Retitled";
            document.ApplicationProperties.Company = "Example Company";
            document.SetCustomDocumentProperty("Region", "South");
            document.Save(propertyEditPath);
        }

        using (ExcelDocument document = ExcelDocument.Load(propertyEditPath)) {
            Assert.Equal("Edited", document.Sheets[0].Cell(2, 1).GetValue<string>());
            Assert.Equal("Retitled", document.Properties.Title);
            Assert.Equal(excelCreator, document.Properties.Creator);
            Assert.Equal(excelCreated, document.Properties.Created!.Value.ToUniversalTime());
            Assert.Equal("Example Company", document.ApplicationProperties.Company);
            Assert.Equal("South", document.CustomProperties["Region"].Text);
            Assert.True(document.RemoveCustomDocumentProperty("Region"));
            document.Save();
        }

        using (ExcelDocument document = ExcelDocument.Load(propertyEditPath)) {
            Assert.Empty(document.CustomProperties);
            Assert.Equal("Retitled", document.Properties.Title);
        }

        using ZipArchive archive = ZipFile.OpenRead(propertyEditPath);
        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XDocument rootRelationships = XDocument.Load(archive.GetEntry("_rels/.rels")!.Open());
        string[] types = rootRelationships.Root!.Elements(relationships + "Relationship")
            .Select(element => (string)element.Attribute("Type")!).ToArray();
        Assert.Single(types, type => type.EndsWith("/metadata/core-properties", StringComparison.Ordinal));
        Assert.Single(types, type => type.EndsWith("/extended-properties", StringComparison.Ordinal));
        Assert.DoesNotContain(types, type => type.EndsWith("/custom-properties", StringComparison.Ordinal));
        Assert.Null(archive.GetEntry("docProps/custom.xml"));
        string contentTypes = new StreamReader(archive.GetEntry("[Content_Types].xml")!.Open()).ReadToEnd();
        Assert.DoesNotContain("custom-properties", contentTypes, StringComparison.Ordinal);
    }

    [Fact]
    public void Xlsb_FailedSaveKeepsPropertyEditsForTheNextSave() {
        string sourcePath = XlsbExcelGeneratedPath();
        string savedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        using (ExcelDocument document = ExcelDocument.Load(sourcePath)) {
            document.Properties.Title = "Retitled";
            document.SetCustomDocumentProperty("Region", "South");
            using (var failing = new XlsbFailingWriteStream()) {
                Assert.Throws<IOException>(() => document.Save(failing, new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
            }
            document.Save(savedPath);
        }

        using ExcelDocument reloaded = ExcelDocument.Load(savedPath);
        Assert.Equal("Retitled", reloaded.Properties.Title);
        Assert.Equal("South", reloaded.CustomProperties["Region"].Text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Xlsb_SaveWritesPropertyPartsEditedThroughTheOpenXmlPackage(bool editCell) {
        string savedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        using (ExcelDocument document = ExcelDocument.Load(XlsbExcelGeneratedPath())) {
            var properties = document.OpenXmlDocument.ExtendedFilePropertiesPart!.Properties!;
            properties.Company = new DocumentFormat.OpenXml.ExtendedProperties.Company("Direct Company");
            if (editCell) document.Sheets[0].CellValue(2, 1, "Edited");
            document.Save(savedPath);
        }

        using ExcelDocument reloaded = ExcelDocument.Load(savedPath);
        Assert.Equal("Direct Company", reloaded.ApplicationProperties.Company);
        Assert.Equal(editCell ? "Edited" : "Alpha", reloaded.Sheets[0].Cell(2, 1).GetValue<string>());
    }

    [Fact]
    public void Xlsb_ReadOnlyLoadReportsTheStoredModifiedTime() {
        var stored = new DateTime(2026, 7, 16, 7, 48, 15, DateTimeKind.Utc);
        string copyPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        using (ExcelDocument readOnly = ExcelDocument.Load(XlsbExcelGeneratedPath(), new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
            Assert.Equal(stored, readOnly.Properties.Modified!.Value.ToUniversalTime());
        }

        using (ExcelDocument writable = ExcelDocument.Load(XlsbExcelGeneratedPath())) {
            Assert.Equal(stored, writable.Properties.Modified!.Value.ToUniversalTime());
            writable.Save(copyPath);
        }

        using ExcelDocument copy = ExcelDocument.Load(copyPath, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        Assert.Equal(stored, copy.Properties.Modified!.Value.ToUniversalTime());
    }

    [Theory]
    [InlineData("custom-root")]
    [InlineData("extended-root")]
    [InlineData("root-relationship")]
    [InlineData("core-target-workbook")]
    public void Xlsb_UnreadablePropertyMetadataIsSkippedAndTheWorkbookStillLoadsAndSaves(string damage) {
        string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        string savedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        File.WriteAllBytes(sourcePath, EditXlsbPackage(File.ReadAllBytes(XlsbExcelGeneratedPath()), archive => {
            switch (damage) {
                case "custom-root":
                    AddCustomPropertiesPart(archive, "<unexpected xmlns=\"urn:example\"/>");
                    break;
                case "extended-root":
                    ReplaceEntry(archive, "docProps/app.xml", "<?xml version=\"1.0\"?><unexpected xmlns=\"urn:example\"/>");
                    break;
                case "core-target-workbook":
                    // A mistyped relationship must never become the write target for the core properties.
                    archive.GetEntry("docProps/core.xml")!.Delete();
                    ReplaceEntry(archive, "_rels/.rels", ReadEntry(archive, "_rels/.rels").Replace(
                        "Target=\"docProps/core.xml\"", "Target=\"xl/workbook.bin\""));
                    break;
                default:
                    string relationships = ReadEntry(archive, "_rels/.rels");
                    ReplaceEntry(archive, "_rels/.rels", relationships.Replace("</Relationships>",
                        "<Relationship Id=\"rIdBroken\" Type=\"urn:example:broken\"/></Relationships>"));
                    break;
            }
        }));

        using (ExcelDocument document = ExcelDocument.Load(sourcePath)) {
            Assert.Contains(document.XlsbImportDiagnostics, diagnostic => diagnostic.Code == "xlsb-document-properties-skipped");
            Assert.Equal("Alpha", document.Sheets[0].Cell(2, 1).GetValue<string>());
            Assert.Empty(document.CustomProperties);
            document.Properties.Title = "Retitled";
            document.Sheets[0].CellValue(3, 1, "Edited");
            document.Save(savedPath);
        }

        using ExcelDocument reloaded = ExcelDocument.Load(savedPath);
        Assert.Equal("Edited", reloaded.Sheets[0].Cell(3, 1).GetValue<string>());
        if (damage == "root-relationship") {
            // The save keeps the package relationships as they were, so the properties stay unreadable.
            Assert.Contains(reloaded.XlsbImportDiagnostics, diagnostic => diagnostic.Code == "xlsb-document-properties-skipped");
        } else {
            Assert.Equal("Retitled", reloaded.Properties.Title);
        }
    }

    [Fact]
    public void Xlsb_RemovingTheLastCustomPropertyKeepsRootMetadata() {
        string sourcePath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsb");
        File.WriteAllBytes(sourcePath, EditXlsbPackage(File.ReadAllBytes(XlsbExcelGeneratedPath()), archive =>
            AddCustomPropertiesPart(archive,
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\" " +
                "xmlns:x=\"urn:example\" x:sourceId=\"7\">" +
                "<property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"Region\"><vt:lpwstr>North</vt:lpwstr></property>" +
                "</Properties>")));

        using (ExcelDocument document = ExcelDocument.Load(sourcePath)) {
            Assert.Equal("North", document.CustomProperties["Region"].Text);
            Assert.True(document.RemoveCustomDocumentProperty("Region"));
            document.Save();
        }

        using (ExcelDocument reloaded = ExcelDocument.Load(sourcePath)) {
            Assert.Empty(reloaded.CustomProperties);
        }
        using ZipArchive saved = ZipFile.OpenRead(sourcePath);
        string custom = ReadEntry(saved, "docProps/custom.xml");
        Assert.Contains("sourceId=\"7\"", custom, StringComparison.Ordinal);
        Assert.DoesNotContain("Region", custom, StringComparison.Ordinal);
    }

    private static byte[] EditXlsbPackage(byte[] package, Action<ZipArchive> edit) {
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) edit(archive);
        return stream.ToArray();
    }

    private static void AddCustomPropertiesPart(ZipArchive archive, string rootXml) {
        ReplaceEntry(archive, "docProps/custom.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" + rootXml);
        ReplaceEntry(archive, "_rels/.rels", ReadEntry(archive, "_rels/.rels").Replace("</Relationships>",
            "<Relationship Id=\"rIdCustom\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties\" Target=\"docProps/custom.xml\"/></Relationships>"));
        ReplaceEntry(archive, "[Content_Types].xml", ReadEntry(archive, "[Content_Types].xml").Replace("</Types>",
            "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/></Types>"));
    }

    private static string ReadEntry(ZipArchive archive, string name) {
        using var reader = new StreamReader(archive.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }

    private static void ReplaceEntry(ZipArchive archive, string name, string content) {
        archive.GetEntry(name)?.Delete();
        using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
        writer.Write(content);
    }

    private static string XlsbExcelGeneratedPath() =>
        Path.Combine(AppContext.BaseDirectory, "Documents", "XlsbCorpus", "excel-generated", "basic-values-formula.xlsb");

    private sealed class XlsbFailingWriteStream : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("Simulated write failure.");
        public override void WriteByte(byte value) => throw new IOException("Simulated write failure.");
    }

    private static void SetXlsbTestProperties(ExcelDocument document) {
        document.Properties.Title = "Quarterly figures";
        document.Properties.Subject = "Finance";
        document.Properties.Creator = "Report Author";
        document.Properties.Keywords = "budget";
        document.Properties.Description = "Figures for the quarter";
        document.Properties.LastModifiedBy = "Report Editor";
        document.Properties.Category = "Reports";
        document.Properties.Created = XlsbPropertiesCreated;
        document.ApplicationProperties.Company = "Example Company";
        document.ApplicationProperties.Manager = "Example Manager";
        document.SetCustomDocumentProperty("Region", "North");
        document.SetCustomDocumentProperty("Revision Count", 7);
    }
}
