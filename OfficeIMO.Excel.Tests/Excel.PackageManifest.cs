using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelPackageManifestTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Promoted_macro_workbook_manifest_matches_the_actual_main_part(bool legacySource) {
        using var document = legacySource
            ? ExcelDocument.Load(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "LegacyXlsCorpus", "openpreserve-format-corpus", "valid.xls"))
            : ExcelDocument.Create();
        if (!legacySource) document.AddWorksheet("Data").CellValue(1, 1, "Preserved");
        var project = OfficeVbaProject.Create(); project.AddModule("Notes", "' Inert manifest regression fixture\r\n");
        document.SetVbaProject(project);
        using var output = new MemoryStream();
        document.Save(output, new ExcelSaveOptions { DisableFastPackageWriter = true, LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.False(OfficeIMO.Excel.Utilities.ExcelPackageUtilities.NormalizeContentTypes(output, leaveOpen: true));
        using (var archive = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true)) {
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
            using var relsStream = archive.GetEntry("_rels/.rels")!.Open();
            string target = XDocument.Load(relsStream).Root!.Elements(relationships + "Relationship")
                .Single(item => ((string?)item.Attribute("Type"))?.EndsWith("/officeDocument", StringComparison.Ordinal) == true).Attribute("Target")!.Value;
            target = "/" + target.TrimStart('/');
            using var typesStream = archive.GetEntry("[Content_Types].xml")!.Open();
            var types = XDocument.Load(typesStream);
            var workbook = types.Root!.Elements(contentTypes + "Override").SingleOrDefault(item => (string?)item.Attribute("PartName") == target);
            Assert.NotNull(workbook);
            Assert.Equal("application/vnd.ms-excel.sheet.macroEnabled.main+xml", (string?)workbook!.Attribute("ContentType"));
            foreach (var part in types.Root.Elements(contentTypes + "Override")) Assert.NotNull(archive.GetEntry(part.Attribute("PartName")!.Value.TrimStart('/')));
        }
        output.Position = 0; using var saved = ExcelDocument.Load(output);
        Assert.True(saved.HasMacros); Assert.Single(saved.ReadVbaProject()!.Modules);
    }
    [Fact]
    public void Template_copy_updates_the_relocated_main_part_without_losing_the_project() {
        using var document = ExcelDocument.Create(); document.AddWorksheet("Data").CellValue(1, 1, "Template value");
        var project = OfficeVbaProject.Create(); project.AddModule("Notes", "' Inert template fixture\r\n"); document.SetVbaProject(project);
        using var source = new MemoryStream(); document.Save(source, new ExcelSaveOptions { DisableFastPackageWriter = true });
        string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-Manifest-" + Guid.NewGuid().ToString("N") + ".xltm");
        try {
            ExcelDocument.CopyPackage(source, path);
            using var copy = ExcelDocument.Load(path);
            Assert.Equal(DocumentFormat.OpenXml.SpreadsheetDocumentType.MacroEnabledTemplate, copy.OpenXmlDocument.DocumentType);
            Assert.Equal("Template value", copy.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.Single(copy.ReadVbaProject()!.Modules);
        } finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void Unrelated_vendor_root_relationships_do_not_interfere_with_Office_parts() {
        using var document = ExcelDocument.Create(); document.AddWorksheet("Data").CellValue(1, 1, "Vendor metadata");
        using var output = new MemoryStream(); document.Save(output, new ExcelSaveOptions { DisableFastPackageWriter = true });
        using (var package = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            var entry = package.GetEntry("_rels/.rels")!;
            XDocument relationships;
            using (var input = entry.Open()) relationships = XDocument.Load(input);
            XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
            int index = 0;
            foreach (string suffix in new[] { "/officeDocument", "/extended-properties", "/metadata/core-properties" })
                relationships.Root!.Add(new XElement(ns + "Relationship", new XAttribute("Id", "vendor" + ++index),
                    new XAttribute("Type", "https://vendor.example/relationships" + suffix), new XAttribute("TargetMode", "External"),
                    new XAttribute("Target", "https://vendor.example/metadata")));
            entry.Delete(); using var replacement = package.CreateEntry("_rels/.rels").Open(); relationships.Save(replacement);
        }
        output.Position = 0;
        using var saved = ExcelDocument.Load(output);
        Assert.Equal("Vendor metadata", saved.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(3, saved.OpenXmlDocument.ExternalRelationships.Count());
        output.Position = 0;
        Assert.False(OfficeIMO.Excel.Utilities.ExcelPackageUtilities.NormalizeContentTypes(output, leaveOpen: true));
    }

}
