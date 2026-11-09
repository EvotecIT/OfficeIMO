using System.Globalization;
using AngleSharp.Html.Parser;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void PowerPointTableBorderLossesKeepCountsAndFirstSeenOrderPerTable() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var repeated = slide.AddTable(1, 2, OfficeLength.Points(20), OfficeLength.Points(20),
            OfficeLength.Points(200), OfficeLength.Points(40));
        foreach (var cell in repeated.Rows[0].Cells) {
            cell.Text = "Repeated";
            cell.SetBorders(PowerPointTableCellBorders.Left | PowerPointTableCellBorders.Top,
                OfficeColor.Black, OfficeLength.Points(1), PowerPointLineDashStyle.DashDot);
            cell.SetBorders(PowerPointTableCellBorders.DiagonalDown, OfficeColor.Black);
        }
        var single = slide.AddTable(1, 1, OfficeLength.Points(20), OfficeLength.Points(80),
            OfficeLength.Points(100), OfficeLength.Points(40));
        single.Rows[0].Cells[0].Text = "Single";
        single.Rows[0].Cells[0].SetBorders(PowerPointTableCellBorders.Bottom,
            OfficeColor.Black, OfficeLength.Points(1), PowerPointLineDashStyle.DashDot);

        string html = document.ToHtml(null, out OfficeConversionReport report);

        var cells = new HtmlParser().ParseDocument(html).QuerySelectorAll("td");
        Assert.Equal(new[] { "Repeated", "Repeated", "Single" }, cells.Select(cell => cell.TextContent));
        Assert.Contains("border-top:1pt dashed", cells[0].GetAttribute("style"));
        Assert.True(report.HasLoss);
        var warnings = report.Diagnostics.Where(diagnostic => diagnostic.Code == "powerpoint-table-border-approximated").ToArray();
        Assert.Equal(3, warnings.Length);
        string repeatedSource = "powerpoint:table:" + repeated.DrawingOrder.ToString(CultureInfo.InvariantCulture);
        Assert.Equal(repeatedSource, warnings[0].Location);
        Assert.Equal("Dash-dot table borders are approximated by CSS dashed borders. (4 occurrences)", warnings[0].Message);
        Assert.Equal(repeatedSource, warnings[1].Location);
        Assert.Equal("Diagonal table borders have no CSS table-border equivalent. (2 occurrences)", warnings[1].Message);
        Assert.Equal("powerpoint:table:" + single.DrawingOrder.ToString(CultureInfo.InvariantCulture), warnings[2].Location);
        Assert.Equal("Dash-dot table borders are approximated by CSS dashed borders.", warnings[2].Message);
        Assert.All(warnings, warning => {
            Assert.Equal(OfficeDiagnosticSeverity.Warning, warning.Severity);
            Assert.Equal(OfficeConversionLossKind.Approximation, warning.LossKind);
        });
    }

    [Fact]
    public void PowerPointSemanticTableKeepsSourceCellPaintAndGridDimensions() {
        using var document = PowerPointDocument.Create();
        var table = document.AddSlide().AddTable(1, 2,
            OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(200), OfficeLength.Points(80));
        table.SetColumnWidths(OfficeLength.Points(80), OfficeLength.Points(120));
        table.SetRowHeight(0, OfficeLength.Points(36));
        var cell = table.Rows[0].Cells[0];
        cell.Text = "Authored";
        cell.FillColor = OfficeColor.Parse("#123456");
        cell.SetBorders(PowerPointTableCellBorders.All, OfficeColor.Parse("#654321"), OfficeLength.Points(2));
        cell.PaddingLeft = OfficeLength.Zero;
        cell.HorizontalAlignment = PowerPointTextAlignment.Right;
        cell.VerticalAlignment = PowerPointTextVerticalAlignment.Bottom;
        var transparent = table.Rows[0].Cells[1];
        transparent.Text = "Unfilled";
        transparent.NoFill = true;
        transparent.HideBorders(PowerPointTableCellBorders.All);

        var html = new HtmlParser().ParseDocument(document.ToHtml());
        var cells = html.QuerySelectorAll("td").ToArray();
        var columns = html.QuerySelectorAll("col").ToArray();

        Assert.Contains("width:80pt", columns[0].GetAttribute("style"));
        Assert.Contains("width:120pt", columns[1].GetAttribute("style"));
        Assert.Contains("height:36pt", html.QuerySelector("tr")!.GetAttribute("style"));
        string style = cells[0].GetAttribute("style")!;
        Assert.Contains("background-color:#123456", style);
        Assert.Contains("border-left:2pt solid #654321", style);
        Assert.Contains("padding-left:0pt", style);
        Assert.Contains("text-align:right", style);
        Assert.Contains("vertical-align:bottom", style);
        Assert.Contains("border-left:none", cells[1].GetAttribute("style"));
        Assert.Contains("background-color:#00000000", cells[1].GetAttribute("style"));
    }

    [Fact]
    public void PowerPointInheritedNoFillKeepsTheSlideBackground() {
        using var document = PowerPointDocument.Create();
        document.SlideSize.SetSize(OfficeLength.Points(180), OfficeLength.Points(100));
        var slide = document.AddSlide();
        slide.BackgroundColor = OfficeColor.Yellow;
        var table = slide.AddTable(1, 1, OfficeLength.Points(20), OfficeLength.Points(20),
            OfficeLength.Points(120), OfficeLength.Points(48));
        table.StyleId = "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}";
        table.FirstRow = true;
        table.BandedRows = false;

        var part = document.OpenXmlDocument.PresentationPart!;
        part.DeletePart(part.TableStylesPart!);
        using var source = new MemoryStream(document.ToBytes(), writable: false);
        using var readOnly = PowerPointDocument.Load(source,
            new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        string sourceTable = readOnly.Slides[0].Tables.Single().TableElement.OuterXml;

        var html = new HtmlParser().ParseDocument(readOnly.ToHtml());
        Assert.Contains("background-color:#00000000", html.QuerySelector("td")!.GetAttribute("style"));
        var png = readOnly.Slides[0].ToImage(OfficeImageExportFormat.Png);
        Assert.True(OfficePngReader.TryDecode(png.Bytes, out var pixels));
        Assert.Equal(OfficeColor.Yellow, pixels!.GetPixel(32, 32));
        Assert.Null(readOnly.OpenXmlDocument.PresentationPart!.TableStylesPart);
        Assert.Equal(sourceTable, readOnly.Slides[0].Tables.Single().TableElement.OuterXml);
    }

    [Fact]
    public void PowerPointInheritedDoubleBordersKeepTwoStrokesAndTheirOverallWidth() {
        using var document = PowerPointDocument.Create();
        var slide = document.AddSlide();
        var table = slide.AddTable(2, 1, OfficeLength.Points(20), OfficeLength.Points(20),
            OfficeLength.Points(120), OfficeLength.Points(60));
        table.StyleId = "{B301B821-A1FF-4177-AEE7-76D212191A09}";
        table.FirstRow = false;
        table.LastRow = true;
        table.BandedRows = false;
        table.SetRowHeights(OfficeLength.Points(30), OfficeLength.Points(30));

        var html = new HtmlParser().ParseDocument(document.ToHtml());
        Assert.Contains("border-top:4pt double", html.QuerySelectorAll("td")[1].GetAttribute("style"));
        var svg = System.Xml.Linq.XElement.Parse(Encoding.UTF8.GetString(slide.ToImage(OfficeImageExportFormat.Svg).Bytes));
        var strokes = svg.Descendants().Where(element => element.Name.LocalName == "line" &&
            double.TryParse(element.Attribute("stroke-width")?.Value, System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture, out double width) && Math.Abs(width - 4D / 3D) < .001D).ToArray();
        Assert.Equal(2, strokes.Length);
        double Y(System.Xml.Linq.XElement line) => double.Parse(line.Attribute("y1")!.Value, System.Globalization.CultureInfo.InvariantCulture);
        Assert.InRange(Math.Abs(Y(strokes[0]) - Y(strokes[1])), 2.66D, 2.67D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeTableShellStylingRequiresAnExplicitProfile(bool powerPoint) {
        string plain;
        string styled;
        if (powerPoint) {
            using var document = PowerPointDocument.Create();
            var table = document.AddSlide().AddTable(2, 2,
                OfficeLength.Points(40), OfficeLength.Points(40), OfficeLength.Points(200), OfficeLength.Points(80));
            table.Rows[0].Cells[0].Text = "Qty";
            table.Rows[1].Cells[0].Text = "12";
            plain = document.ToHtml();
            styled = document.ToHtml(PowerPointHtmlExportOptions.CreateSemanticSlidesProfile());
        } else {
            using var document = ExcelDocument.Create(new MemoryStream());
            var sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("Qty");
            sheet.Cell(2, 1).SetValue(12);
            plain = document.ToHtml();
            styled = document.ToHtml(ExcelHtmlExportOptions.CreateSemanticTablesProfile());
        }

        var parser = new HtmlParser();
        var plainDocument = parser.ParseDocument(plain);
        var styledDocument = parser.ParseDocument(styled);
        Assert.NotNull(plainDocument.QuerySelector("table"));
        Assert.Contains("12", plainDocument.Body!.TextContent);
        Assert.Empty(plainDocument.Head!.QuerySelectorAll("style"));
        Assert.NotEmpty(styledDocument.Head!.QuerySelectorAll("style"));
    }
}
