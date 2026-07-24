using System.IO.Compression;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelAllSeverityBatch15SecurityTests {
    [Fact]
    public void FastSheetReaderIgnoresLookalikeElementsAndWrongRelationshipNamespaces() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-sheet-fast-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                document.AddWorksheet("Data").CellValue(1, 1, "safe");
                document.Save();
            }

            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                ZipArchiveEntry entry = archive.GetEntry("xl/workbook.xml")!;
                XDocument workbook;
                using (Stream input = entry.Open()) workbook = XDocument.Load(input);
                XNamespace spreadsheet = workbook.Root!.Name.Namespace;
                XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace extension = "urn:attacker-extension";
                XElement sheets = workbook.Root.Element(spreadsheet + "sheets")!;
                string realRelationshipId = sheets.Elements(spreadsheet + "sheet").First().Attribute(relationships + "id")!.Value;

                workbook.Root.AddFirst(new XElement(extension + "sheets",
                    new XElement(extension + "sheet",
                        new XAttribute("name", "ExtensionBogus"),
                        new XAttribute(relationships + "id", realRelationshipId))));
                sheets.AddFirst(new XElement(spreadsheet + "sheet",
                    new XAttribute("name", "WrongNamespaceBogus"),
                    new XAttribute(XNamespace.Xmlns + "r", "urn:attacker-relationships"),
                    new XAttribute(XName.Get("id", "urn:attacker-relationships"), realRelationshipId)));

                entry.Delete();
                ZipArchiveEntry replacement = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
                using Stream output = replacement.Open();
                workbook.Save(output);
            }

            using ExcelDocumentReader reader = ExcelDocumentReader.Open(path);
            Assert.Equal(new[] { "Data" }, reader.GetSheetNames());
            Assert.Equal(1, reader.SheetCount);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void UncachedSheetReadsDoNotMutateSharedSheetIdState() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        document.AddWorksheet("Data");
        document.SheetCachingEnabled = false;

        Parallel.For(0, 200, _ => Assert.Single(document.Sheets));
    }

    [Fact]
    public void TableTotalsRetainLastCaseInsensitiveDuplicate() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(1, 2, "Value");
        sheet.CellValue(2, 1, "A");
        sheet.CellValue(2, 2, 1);
        sheet.AddTable("A1:B2", true, "DataTable", OfficeIMO.Excel.TableStyle.TableStyleMedium2);
        var totals = new Dictionary<string, TotalsRowFunctionValues>(StringComparer.Ordinal) {
            ["Name"] = TotalsRowFunctionValues.Count,
            ["NAME"] = TotalsRowFunctionValues.None
        };

        Exception? exception = Record.Exception(() => sheet.SetTableTotals("A1:B2", totals));

        Assert.Null(exception);
    }

    [Fact]
    public void SummarizeOverflowFitsMoreIntoSingleAvailableColumn() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-summary-one-column-" + Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using (ExcelDocument document = ExcelDocument.Create(path)) {
                document.Compose("Report", composer => {
                    composer.Columns(2, columns => columns[0].TableFrom(
                        new[] { new WideRow("Alpha", 1, 2) }, title: "Wide"),
                        columnWidth: 1,
                        overflow: OverflowMode.Summarize);
                    composer.Finish(autoFitColumns: false);
                });
                document.Save();
            }

            using SpreadsheetDocument package = SpreadsheetDocument.Open(path, false);
            Table table = Assert.Single(package.WorkbookPart!.WorksheetParts.First().TableDefinitionParts).Table!;
            Assert.Equal("A2:A3", table.Reference!.Value);
            Assert.Equal("More", Assert.Single(table.TableColumns!.Elements<TableColumn>()).Name!.Value);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ParallelRowAutoFitRequestsSerializeOpenXmlTraversal() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        for (int row = 1; row <= 100; row++) sheet.CellValue(row, 1, "row " + row);

        Parallel.Invoke(
            () => sheet.AutoFitRows(ExecutionMode.Parallel),
            () => sheet.AutoFitRows(ExecutionMode.Parallel),
            () => sheet.AutoFitRows(ExecutionMode.Parallel));

        Assert.NotNull(sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First().Height);
    }

    private sealed record WideRow(string Name, int Score, int Total);
}
