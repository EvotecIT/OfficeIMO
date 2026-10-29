using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeChartOptionsAuthoringTests {
    private static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace SpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace WordDrawing = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly Trend[] Rows = { new("First", 12.5m), new("Next", 19.75m) };
    private static readonly OfficeChartData Data = new(new[] { "First", "Next" },
        new[] { new OfficeSeries("Revenue", new[] { 12.5, 19.75 }) });

    [Theory]
    [InlineData("Word", "shared")]
    [InlineData("Word", "typed")]
    [InlineData("Word", "named")]
    [InlineData("WordParagraph", "shared")]
    [InlineData("Excel", "shared")]
    [InlineData("Excel", "typed")]
    [InlineData("Excel", "named")]
    [InlineData("PowerPoint", "shared")]
    [InlineData("PowerPoint", "typed")]
    [InlineData("PowerPoint", "named")]
    public void SharedOptionsWriteTitlesAndExactPhysicalExtents(string format, string source) {
        var options = new OfficeChartOptions {
            Title = "Quarterly revenue", Width = OfficeLength.Emus(3_600_001), Height = OfficeLength.Emus(1_800_003)
        };
        byte[] bytes = Create(format, source, options);
        var documents = ReadXmlParts(bytes);
        var chart = Assert.Single(documents, xml => xml.Root?.Name == Chart + "chartSpace");
        Assert.Equal(options.Title, string.Concat(chart.Descendants(Chart + "title").Descendants(Drawing + "t").Select(node => node.Value)));
        Assert.Equal(new[] { "12.5", "19.75" }, chart.Descendants(Chart + "numCache").Elements(Chart + "pt")
            .Elements(Chart + "v").Select(node => node.Value));
        var extents = documents.SelectMany(xml => xml.Descendants(format.StartsWith("Word", StringComparison.Ordinal)
            ? WordDrawing + "extent" : format == "Excel" ? SpreadsheetDrawing + "ext" : Drawing + "ext"));
        Assert.Contains(extents, node => (string?)node.Attribute("cx") == "3600001" && (string?)node.Attribute("cy") == "1800003");
        if (format == "Excel") {
            using var workbook = ExcelDocument.Load(new MemoryStream(bytes));
            var sheet = workbook.Sheets.Single(item => item.Name == "Trend");
            var table = sheet.Table("TrendData");
            Assert.Equal(Rows, table.RowsAs<Trend>().ToArray());
            Assert.Single(sheet.Charts);
        }
    }

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void FormatOptionsRetainFrameAnchorAndAccessibility(string format) {
        OfficeChartOptions options = format switch {
            "Word" => new WordChartOptions { RoundedCorners = true },
            "Excel" => new ExcelChartOptions { Row = 4, Column = 7 },
            _ => new PowerPointChartOptions {
                Left = OfficeLength.Emus(-12_345), Top = OfficeLength.Emus(23_457),
                Accessibility = new PowerPointChartAccessibilityOptions { Name = "Revenue chart", AlternativeText = "Two weekly observations" }
            }
        };
        options.Title = "Placed chart";
        var documents = ReadXmlParts(Create(format, "shared", options));
        switch (format) {
            case "Word":
                Assert.Contains(documents.SelectMany(xml => xml.Descendants(Chart + "roundedCorners")),
                    node => (string?)node.Attribute("val") == "1");
                break;
            case "Excel":
                var origin = Assert.Single(documents.SelectMany(xml => xml.Descendants(SpreadsheetDrawing + "from")));
                Assert.Equal("3", origin.Element(SpreadsheetDrawing + "row")?.Value);
                Assert.Equal("6", origin.Element(SpreadsheetDrawing + "col")?.Value);
                break;
            default:
                Assert.Contains(documents.SelectMany(xml => xml.Descendants(Drawing + "off")),
                    node => (string?)node.Attribute("x") == "-12345" && (string?)node.Attribute("y") == "23457");
                Assert.Contains(documents.SelectMany(xml => xml.Descendants()),
                    node => (string?)node.Attribute("name") == "Revenue chart" && ((string?)node.Attribute("descr"))?.StartsWith("Two weekly observations", StringComparison.Ordinal) == true
                        && ((string?)node.Attribute("descr"))?.Contains("Data summary:", StringComparison.Ordinal) == true);
                break;
        }
    }

    [Theory]
    [InlineData("Word")]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void InvalidPhysicalSizeLeavesExistingContentUsable(string format) {
        var options = new OfficeChartOptions { Width = OfficeLength.Emus(0) };
        using var output = new MemoryStream();
        if (format == "Word") {
            using var document = WordDocument.Create();
            document.AddParagraph("Keep this paragraph");
            Assert.Throws<ArgumentOutOfRangeException>(() => document.AddChart(OfficeChartType.Column, Data, options));
            document.Save(output);
        } else if (format == "Excel") {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Trend");
            sheet.AddTable(Rows, new OfficeTableOptions { Name = "TrendData" });
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.AddChart(OfficeChartType.Column, Data, options));
            Assert.Equal(Rows, sheet.Table("TrendData").RowsAs<Trend>().ToArray());
            document.Save(output);
        } else {
            using var document = PowerPointDocument.Create();
            var slide = document.AddSlide();
            Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddChart(OfficeChartType.Column, Data, options));
            document.Save(output);
        }
        Assert.DoesNotContain(ReadXmlParts(output.ToArray()), xml => xml.Root?.Name == Chart + "chartSpace");
    }

    private static byte[] Create(string format, string source, OfficeChartOptions options) {
        using var output = new MemoryStream();
        if (format.StartsWith("Word", StringComparison.Ordinal)) {
            using var document = WordDocument.Create();
            if (format == "WordParagraph") document.AddParagraph("Lead in").AddChart(OfficeChartType.Column, Data, options);
            else if (source == "typed") document.AddChart(OfficeChartType.Column, Rows, row => row.Week, options, OfficeSeries.From(Rows, row => row.Revenue));
            else if (source == "named") document.AddChart(OfficeChartType.Column, Rows, nameof(Trend.Week), options, nameof(Trend.Revenue));
            else document.AddChart(OfficeChartType.Column, Data, options);
            document.Save(output);
        } else if (format == "Excel") {
            using var document = ExcelDocument.Create();
            var sheet = document.AddSheet("Trend");
            sheet.AddTable(Rows, new OfficeTableOptions { Name = "TrendData" });
            if (source == "typed") sheet.AddChart(OfficeChartType.Column, Rows, row => row.Week, options, OfficeSeries.From(Rows, row => row.Revenue));
            else if (source == "named") sheet.AddChart(OfficeChartType.Column, Rows, nameof(Trend.Week), options, nameof(Trend.Revenue));
            else sheet.AddChart(OfficeChartType.Column, Data, options);
            document.Save(output);
        } else {
            using var document = PowerPointDocument.Create();
            var slide = document.AddSlide();
            if (source == "typed") slide.AddChart(OfficeChartType.Column, Rows, row => row.Week, options, OfficeSeries.From(Rows, row => row.Revenue));
            else if (source == "named") slide.AddChart(OfficeChartType.Column, Rows, nameof(Trend.Week), options, nameof(Trend.Revenue));
            else slide.AddChart(OfficeChartType.Column, Data, options);
            document.Save(output);
        }
        return output.ToArray();
    }

    private static List<XDocument> ReadXmlParts(byte[] bytes) {
        using var archive = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
        var documents = new List<XDocument>();
        foreach (var entry in archive.Entries.Where(entry => entry.FullName.EndsWith(".xml", StringComparison.Ordinal))) {
            using var stream = entry.Open();
            documents.Add(XDocument.Load(stream));
        }
        return documents;
    }

    public sealed record Trend(string Week, decimal Revenue);
}
