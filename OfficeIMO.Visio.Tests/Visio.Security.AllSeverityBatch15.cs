using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioAllSeverityBatch15SecurityTests {
    [Theory]
    [InlineData("1E309")]
    [InlineData("NaN")]
    [InlineData("2147483648")]
    public void LoadIgnoresOutOfRangePatternValues(string value) {
        string path = CreateSample();
        try {
            UpdateXml(path, "visio/pages/page1.xml", document => {
                XElement shape = document.Descendants().First(element => element.Name.LocalName == "Shape");
                shape.Add(new XElement(shape.Name.Namespace + "Cell",
                    new XAttribute("N", "LinePattern"),
                    new XAttribute("V", value)));
            });

            Exception? exception = Record.Exception(() => VisioDocument.Load(path));

            Assert.Null(exception);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void StreamingValidatorReportsDuplicateRelationshipIdsWithoutThrowing() {
        string path = CreateSample();
        try {
            UpdateXml(path, "visio/pages/_rels/pages.xml.rels", document => {
                XElement relationship = document.Root!.Elements().First();
                document.Root.Add(new XElement(relationship));
            });
            var validator = new VsdxPackageValidator();

            bool valid = validator.ValidateFileStreaming(path);

            Assert.False(valid);
            Assert.Contains(validator.Errors, error => error.Contains("Duplicate relationship Id", StringComparison.Ordinal));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static string CreateSample() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-visio-b15-" + Guid.NewGuid().ToString("N") + ".vsdx");
        VisioDocument document = VisioDocument.Create(path);
        VisioPage page = document.AddPage("Page-1", 8.5, 11);
        page.Shapes.Add(new VisioShape("1", 1, 1, 2, 1, "Sample"));
        document.Save();
        return path;
    }

    private static void UpdateXml(string path, string entryName, Action<XDocument> update) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(entryName)!;
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input);
        update(document);
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output);
    }
}
