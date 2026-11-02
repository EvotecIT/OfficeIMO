using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeTypedChartAuthoringTests {
    [Theory]
    [InlineData("Word", false)]
    [InlineData("Word", true)]
    [InlineData("Excel", false)]
    [InlineData("Excel", true)]
    [InlineData("PowerPoint", false)]
    [InlineData("PowerPoint", true)]
    public void TypedAndNamedSelectorsProduceEditableNativeCharts(string format, bool named) {
        var rows = new[] { new Trend { Week = "First", Revenue = 12.5m }, new Trend { Week = "Next", Revenue = 19.75m } };
        using var output = new MemoryStream();
        switch (format) {
            case "Word":
                using (var document = WordDocument.Create()) {
                    if (named) document.AddChart(OfficeChartType.Column,rows,nameof(Trend.Week),nameof(Trend.Revenue));
                    else document.AddChart(OfficeChartType.Column,rows,row => row.Week,OfficeSeries.From(rows, row => row.Revenue));
                    document.Save(output);
                }
                break;
            case "Excel":
                using (var document = ExcelDocument.Create()) {
                    var sheet = document.AddSheet("Trend");
                    if (named) sheet.AddChart(OfficeChartType.Column, rows, nameof(Trend.Week), nameof(Trend.Revenue));
                    else sheet.AddChart(OfficeChartType.Column, rows, row => row.Week, OfficeSeries.From(rows, row => row.Revenue));
                    document.Save(output);
                }
                break;
            default:
                using (var document = PowerPointDocument.Create()) {
                    var slide = document.AddSlide();
                    if (named) slide.AddChart(OfficeChartType.Column, rows, nameof(Trend.Week), nameof(Trend.Revenue));
                    else slide.AddChart(OfficeChartType.Column, rows, row => row.Week, OfficeSeries.From(rows, row => row.Revenue));
                    document.Save(output);
                }
                break;
        }
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var chartParts = new List<XDocument>();
        foreach (var part in archive.Entries.Where(entry => entry.FullName.Contains("/charts/") && entry.Name.EndsWith(".xml", StringComparison.Ordinal))) {
            using var stream = part.Open();
            var candidate = XDocument.Load(stream);
            if (candidate.Root?.Name == c + "chartSpace") chartParts.Add(candidate);
        }
        var xml = Assert.Single(chartParts);
        Assert.Equal(new[] { "12.5", "19.75" }, xml.Descendants(c + "numCache").Elements(c + "pt").Elements(c + "v").Select(value => value.Value));
        Assert.Contains("Revenue", xml.Descendants(c + "v").Select(value => value.Value));
        if (format == "Excel") {
            Assert.NotEmpty(xml.Descendants(c + "strRef").Elements(c + "f"));
            using var workbook = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            var workbookPart = workbook.WorkbookPart!;
            var text = workbookPart.WorksheetParts.SelectMany(part => part.Worksheet!.Descendants<InlineString>().Select(item => item.InnerText))
                .Concat(workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(item => item.InnerText) ?? Enumerable.Empty<string>());
            Assert.Contains("First", text);
        } else Assert.Contains("First", xml.Descendants(c + "v").Select(value => value.Value));
    }

    [Theory]
    [InlineData(OfficeChartType.Funnel)]
    [InlineData(OfficeChartType.Waterfall)]
    [InlineData(OfficeChartType.BoxWhisker)]
    [InlineData(OfficeChartType.Treemap)]
    [InlineData(OfficeChartType.Sunburst)]
    [InlineData(OfficeChartType.Bubble)]
    public void UnsupportedGeneralExcelChartTypesLeaveTheWorkbookUntouched(OfficeChartType type) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Trend");
        var data = new ExcelChartData(new[] { "First", "Next" }, new[] { new ExcelChartSeries("Revenue", new[] { 1d, 2d }) });
        Assert.Throws<NotSupportedException>(() => sheet.AddChart(type, data, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 }));
        Assert.Single(document.Sheets);
        Assert.Empty(document.Inspect().Worksheets.Single().Cells);
        using var output = new MemoryStream();
        document.Save(output);
        using var archive = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.StartsWith("xl/charts/", StringComparison.Ordinal)
            || entry.FullName.StartsWith("xl/drawings/", StringComparison.Ordinal));
    }

    public sealed class Trend {
        public string Week { get; set; } = "";
        public decimal Revenue { get; set; }
    }
}
