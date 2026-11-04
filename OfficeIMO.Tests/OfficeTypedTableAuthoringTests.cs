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
                    var sheet = document.AddSheet("Products");
                    var table = sheet.AddTable(rows, new OfficeTableOptions { Name = "Products" }, row: 2, column: 3);
                    Assert.Equal("C2:D4", table.Range);
                    table.SetStyle(ExcelTableStyle.TableStyleMedium9);
                    document.Save(output);
                }
                break;
            default:
                using (var document = PowerPointDocument.Create()) {
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
        var sheet = document.AddSheet("Data");
        using var data = new System.Data.DataTable("Input");
        data.Columns.Add("Amount", typeof(decimal));
        data.Rows.Add(12.5m);
        var table = sheet.AddTable(data);
        Assert.Equal("A1:A2", table.Range);
        Assert.Same(sheet, table.Sheet);
        var rangeTable = sheet.AddTable("C1:D3", true, "Second", ExcelTableStyle.TableStyleMedium2);
        Assert.Equal("C1:D3", rangeTable.Range);
    }

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void SharedProjectionOptionsSelectFormatAndAlignNativeColumns(string format) {
        var options = new OfficeTableOptions();
        options.Include.UnionWith(new[] { nameof(Record.Count), nameof(Record.Product) });
        options.Order.Add(nameof(Record.Count));
        options.HeaderTransform = name => "Display " + name;
        options.HeaderRenames[nameof(Record.Count)] = "Total";
        options.Formatters[nameof(Record.Count)] = value => $"Units: {value}";
        options.Alignments.Add(OfficeAlignment.Right);
        var rows = new[] { new Record { Product = "First", Count = 3 } };
        using var output = new MemoryStream();
        switch (format) {
            case "Word":
                using (var document = WordDocument.Create()) {
                    var table = document.AddTable(rows, options);
                    Assert.Equal("Total", table.Rows[0].Cells[0].Paragraphs[0].Text);
                    Assert.Equal(WordParagraphAlignment.Right, table.Rows[1].Cells[0].Paragraphs[0].ParagraphAlignment);
                    document.Save(output);
                }
                break;
            case "Excel":
                using (var document = ExcelDocument.Create()) {
                    var table = document.AddSheet("Data").AddTable(rows, options);
                    Assert.Equal("A1:B2", table.Range);
                    document.Save(output);
                }
                break;
            default:
                using (var document = PowerPointDocument.Create()) {
                    var table = document.AddSlide().AddTable(rows, options);
                    Assert.Equal("Total", table.GetCell(0, 0).Text);
                    Assert.Equal(PowerPointTextAlignment.Right, table.GetCell(1, 0).HorizontalAlignment);
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
        Assert.Contains("Total", text);
        Assert.Contains("Display Product", text);
        Assert.Contains("Units: 3", text);
    }

    [Fact]
    public void CommonExcelOptionsDoNotChangeReusableAdvancedProjection() {
        var projection = new OfficeIMO.Data.ObjectFlattenerOptions();
        projection.Formatters[nameof(Record.Count)] = _ => "Original";
        var options = new ExcelTableOptions { Projection = projection };
        options.Formatters[nameof(Record.Count)] = _ => "Table";
        using var document = ExcelDocument.Create();
        document.AddSheet("Data").AddTable(new[] { new Record { Product = "First", Count = 3 } }, options);
        Assert.Equal("Original", projection.Formatters[nameof(Record.Count)](3));
    }

    [Fact]
    public void ExcelDataTableChecksNativeBoundsBeforeFormattingSelectedValues() {
        using var data = new System.Data.DataTable("Data");
        data.Columns.Add("Name", typeof(string));
        data.Columns.Add("Amount", typeof(int));
        data.Rows.Add("First", 3);
        var options = new OfficeTableOptions();
        options.Formatters["Amount"] = _ => throw new InvalidOperationException("A rejected table must not read values.");
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        Assert.Throws<ArgumentException>(() => sheet.AddTable(data, options, row: 1048576));
        Assert.Throws<ArgumentException>(() => sheet.AddTable(data, options, column: 16384));
        options.Formatters.Clear();
        options.Include.Add("Name");
        var table = sheet.AddTable(data, options, column: 16384);
        Assert.Equal("XFD1:XFD2", table.Range);
    }

    public sealed class Record {
        public string Product { get; set; } = "";
        public int Count { get; set; }
    }
}
