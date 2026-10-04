using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Data;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelDataProjectionContractsTests {
    [Fact]
    public void ObjectProjectionIncludesOnlyReadableNonIndexedPublicProperties() {
        var flattener = new ObjectFlattener();
        var options = new ObjectFlattenerOptions();

        Dictionary<string, object?> values = flattener.Flatten(new IndexedRow(), options);
        Assert.Equal(new[] { "Id" }, flattener.GetPaths(typeof(IndexedRow), options));
        Assert.Equal(7, Assert.Single(values).Value);
    }

    [Theory]
    [InlineData(null, "Metrics.Score")]
    [InlineData("Custom", "Custom.Score")]
    [InlineData("", "Score")]
    public void CollectionHeaderPrefixChangesDisplayedHeaderWithoutChangingSelectionOrFormatting(string? prefix, string expectedHeader) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
        var options = new ObjectFlattenerOptions {
            Columns = new[] { "Metrics.Score" }
        };
        options.CollectionMapColumns["Metrics"] = new CollectionColumnMapping { HeaderPrefix = prefix };
        options.Formatters["Metrics.Score"] = value => (int)value! + 1;
        var row = new MetricsRow();

        Assert.Equal(13, new ObjectFlattener().Flatten(row, options)["Metrics.Score"]);
        try {
            using (var document = ExcelDocument.Create(path)) {
                document.AsFluent().Sheet("Metrics", sheet => sheet.RowsFrom(new[] { row }, settings => {
                    settings.Columns = options.Columns;
                    settings.CollectionMapColumns["Metrics"] = options.CollectionMapColumns["Metrics"];
                    settings.Formatters["Metrics.Score"] = options.Formatters["Metrics.Score"];
                })).End().Save();
            }
            using var package = SpreadsheetDocument.Open(path, false);
            var worksheet = package.WorkbookPart!.WorksheetParts.Single().Worksheet;
            Cell header = worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A1");
            string text = header.DataType?.Value == CellValues.SharedString
                ? package.WorkbookPart.SharedStringTablePart!.SharedStringTable.ChildElements[int.Parse(header.CellValue!.Text)].InnerText
                : header.InlineString?.InnerText ?? header.CellValue?.Text ?? string.Empty;
            Assert.Equal(expectedHeader, text);
            Assert.Equal("13", worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "A2").CellValue!.Text);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class IndexedRow {
        public int Id => 7;
        public int this[int index] => throw new InvalidOperationException("An indexer is not a table column.");
        public int WriteOnly { set => throw new InvalidOperationException("A setter is not a readable column."); }
        public int PrivateGetter { private get; set; }
    }

    private sealed class MetricsRow {
        public Metric[] Metrics { get; } = new[] { new Metric() };
    }

    private sealed class Metric {
        public string Name => "Score";
        public int Value => 12;
    }
}
