using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeTypedTableAuthoringTests {
    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void TypedRowsProduceEditableTablesWithHeaders(string format) {
        var rows = new[] { new Record { Product = "First", Count = 3 }, new Record { Product = "Next", Count = 8 } };
        using var output = new MemoryStream();
        switch (format) {
            case "Word":
                using (var document = WordDocument.Create()) {
                    var table = document.AddTable(rows, new OfficeTableOptions { Name = "Products" });
                    Assert.Equal(3, table.Rows.Count);
                    Assert.Equal("Products", table.Title);
                    document.Save(output);
                }
                break;
            case "Excel":
                using (var document = ExcelDocument.Create()) {
                    var sheet = document.AddWorksheet("Products");
                    var table = sheet.AddTable(rows, new OfficeTableOptions { Name = "Products" }, row: 2, column: 3);
                    Assert.Equal("C2:D4", table.Range);
                    table.SetStyle(ExcelTableStyle.TableStyleMedium9);
                    document.Save(output);
                }
                break;
            default:
                using (var document = PowerPointPresentation.Create()) {
                    var table = document.AddSlide().AddTable(rows, new OfficeTableOptions { Name = "Products" });
                    Assert.Equal(3, table.Rows);
                    Assert.Equal("Products", table.Name);
                    document.Save(output);
                }
                break;
        }
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()));
        var text = new List<string>();
        foreach (var entry in archive.Entries.Where(entry => entry.Name.EndsWith(".xml", StringComparison.Ordinal))) {
            using var stream = entry.Open();
            text.AddRange(XDocument.Load(stream).Descendants().Where(element => !element.HasElements).Select(element => element.Value));
        }
        Assert.Contains("Product", text);
        Assert.Contains("Count", text);
        Assert.Contains("First", text);
        Assert.Contains("Next", text);
    }

    [Fact]
    public void ExcelDataTableRetainsValuesAndReturnsTheCreatedTable() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddWorksheet("Data");
        using var data = new System.Data.DataTable("Input");
        data.Columns.Add("Amount", typeof(decimal));
        data.Rows.Add(12.5m);
        var table = sheet.AddTable(data);
        Assert.Equal("A1:A2", table.Range);
        Assert.Same(sheet, table.Sheet);
        var rangeTable = sheet.AddTable("C1:D3", true, "Second", ExcelTableStyle.TableStyleMedium2);
        Assert.Equal("C1:D3", rangeTable.Range);
    }

    public sealed class Record {
        public string Product { get; set; } = "";
        public int Count { get; set; }
    }
}
