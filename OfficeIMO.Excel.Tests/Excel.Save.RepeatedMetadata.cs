using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelRepeatedMetadataSaveTests {
    private const string CoreContentType = "application/vnd.openxmlformats-package.core-properties+xml";
    private const string CoreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Repeated_extended_file_saves_keep_one_valid_core_properties_part(bool failFirstCommit) {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-RepeatedMetadata-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var document = ExcelDocument.Create();
            ExcelSheet retainedSheet = document.AddSheet("Report");
            retainedSheet.CellValue(1, 1, 42);
            retainedSheet.AddChart(new ExcelChartData(new[] { "Value" },
                    new[] { new ExcelChartSeries("Series", new[] { 42d }) }),
                row: 4, column: 1, width: OfficeLength.Pixels(300), height: OfficeLength.Pixels(180),
                type: OfficeChartType.ColumnClustered, title: "Metadata chart");
            string outputPath = Path.Combine(directory, "report.xlsx");

            if (failFirstCommit) {
                string blockedPath = Path.Combine(directory, "blocked.xlsx");
                Directory.CreateDirectory(blockedPath);
                Exception? failure = Record.Exception(() => document.Save(blockedPath));
                Assert.NotNull(failure);
                Assert.True(failure is IOException || failure is UnauthorizedAccessException, failure.ToString());
                Assert.True(Directory.Exists(blockedPath));
                Assert.False(File.Exists(blockedPath));
            } else {
                document.Save(outputPath);
                Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
                AssertRepeatedMetadataPackage(outputPath, 42);
            }

            // Keep package metadata clean so the next save exercises the same extended writer.
            retainedSheet.CellValue(1, 1, 84);
            document.Save(outputPath);
            Assert.Equal(ExcelSavePackageWriter.ExtendedPackage, document.LastSaveDiagnostics.Writer);
            AssertRepeatedMetadataPackage(outputPath, 84);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void AssertRepeatedMetadataPackage(string path, int expectedCellValue) {
        using (var archive = ZipFile.OpenRead(path)) {
            Assert.Single(archive.Entries, entry => entry.FullName == "docProps/core.xml");
            XNamespace contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types";
            using (Stream stream = archive.GetEntry("[Content_Types].xml")!.Open()) {
                XDocument manifest = XDocument.Load(stream);
                XElement coreOverride = Assert.Single(manifest.Root!.Elements(contentTypes + "Override"),
                    element => (string?)element.Attribute("PartName") == "/docProps/core.xml");
                Assert.Equal(CoreContentType, (string?)coreOverride.Attribute("ContentType"));
            }

            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            using (Stream stream = archive.GetEntry("_rels/.rels")!.Open()) {
                XDocument rootRelationships = XDocument.Load(stream);
                XElement coreRelationship = Assert.Single(rootRelationships.Root!.Elements(relationships + "Relationship"),
                    element => (string?)element.Attribute("Type") == CoreRelationshipType);
                Assert.Equal("docProps/core.xml", (string?)coreRelationship.Attribute("Target"));
            }
        }

        using (var spreadsheet = SpreadsheetDocument.Open(path, false)) {
            Assert.NotNull(spreadsheet.CoreFilePropertiesPart);
            Assert.Equal(CoreContentType, spreadsheet.CoreFilePropertiesPart!.ContentType);
            Assert.NotNull(spreadsheet.PackageProperties.Modified);
            Sheet sheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(item => item.Name == "Report");
            WorksheetPart worksheet = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(sheet.Id!.Value!);
            Cell cell = Assert.Single(worksheet.Worksheet.Descendants<Cell>(), value => value.CellReference?.Value == "A1");
            Assert.Equal(expectedCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture), cell.CellValue?.Text);
            Assert.Single(worksheet.DrawingsPart!.ChartParts);
        }

        using var reopened = ExcelDocument.Load(path);
        Assert.Equal(expectedCellValue, reopened["Report"].CellAt(1, 1).GetValue<int>());
        Assert.Equal("Metadata chart", Assert.Single(reopened["Report"].Charts).Title);
    }
}
