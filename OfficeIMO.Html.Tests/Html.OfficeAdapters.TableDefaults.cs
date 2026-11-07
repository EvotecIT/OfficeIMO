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
