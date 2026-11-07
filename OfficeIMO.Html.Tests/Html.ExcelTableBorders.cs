using AngleSharp.Html.Parser;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Drawing;
using System.Globalization;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlExcelTableBorders {
    [Fact]
    public void SemanticTablesPreserveAuthoredBorderEdgesWithoutAddingAGrid() {
        OfficeTheme theme = OfficeTheme.Default.WithPalette(OfficeTheme.Default.Palette.With(
            accent1: OfficeColor.Parse("123456"), accent2: OfficeColor.Parse("654321")));
        using ExcelDocument workbook = ExcelDocument.Create(new ExcelCreateOptions { Theme = theme });
        ExcelSheet sheet = workbook.AddSheet("Borders");
        sheet.Cell(1, 1).SetValue("Authored edges")
            .SetBorder(ExcelCellBorders.Top, ExcelBorderStyle.Thin, OfficeColor.Accent1)
            .SetBorder(ExcelCellBorders.Right, ExcelBorderStyle.MediumDashed, OfficeColor.Parse("ABCDEF"))
            .SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Double, OfficeColor.Accent2)
            .SetBorder(ExcelCellBorders.Left, ExcelBorderStyle.Dotted, OfficeColor.Parse("0A0B0C"));
        sheet.Cell(1, 2).SetValue("Unbordered");
        sheet.Cell(2, 1).SetValue("Cleared edge").SetBorder(ExcelCellBorders.Top, ExcelBorderStyle.None);

        string html = workbook.ToHtml();
        var document = new HtmlParser().ParseDocument(html);
        var authored = document.QuerySelector("[data-officeimo-cell='A1']")!;
        string authoredStyle = authored.GetAttribute("style")!;
        Assert.Contains("border-top:1px solid #123456", authoredStyle, StringComparison.Ordinal);
        Assert.Contains("border-right:2px dashed #ABCDEF", authoredStyle, StringComparison.Ordinal);
        Assert.Contains("border-bottom:3px double #654321", authoredStyle, StringComparison.Ordinal);
        Assert.Contains("border-left:1px dotted #0A0B0C", authoredStyle, StringComparison.Ordinal);
        string clearedStyle = document.QuerySelector("[data-officeimo-cell='A2']")!.GetAttribute("style")!;
        Assert.Contains("border-top:none", clearedStyle, StringComparison.Ordinal);
        string plainStyle = document.QuerySelector("[data-officeimo-cell='B1']")!.GetAttribute("style") ?? string.Empty;
        Assert.DoesNotContain("border", plainStyle, StringComparison.Ordinal);
        Assert.Contains("border-collapse:collapse", document.QuerySelector("table.officeimo-table")!.GetAttribute("style")!, StringComparison.Ordinal);
    }

    [Fact]
    public void MergedPerimeterBordersReachVisualAndSemanticExports() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Merged");
        sheet.Cell(1, 1).SetValue(42.5).SetNumberFormat("0.00").SetBold().SetFillColor(OfficeColor.Parse("EEDDCC"));
        sheet.Range("A1:B2").Merge();
        sheet.SetColumnWidth(1, ExcelColumnWidth.Length(OfficeLength.Pixels(60)));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Length(OfficeLength.Pixels(60)));
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeLength.Points(15)));
        sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeLength.Points(15)));
        sheet.Range("B1:B2").SetBorder(ExcelCellBorders.Right, ExcelBorderStyle.Medium, OfficeColor.Parse("CC1122"));
        sheet.Range("A2:B2").SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Thick, OfficeColor.Parse("2244CC"));

        string svg = Encoding.UTF8.GetString(sheet.Range("A1:B2").ExportImage(OfficeImageExportFormat.Svg,
            new ExcelImageExportOptions { ShowGridlines = false }).Bytes);
        XElement image = XElement.Parse(svg);
        XElement right = Assert.Single(image.Descendants(), edge => (string?)edge.Attribute("stroke") == "#CC1122");
        XElement bottom = Assert.Single(image.Descendants(), edge => (string?)edge.Attribute("stroke") == "#2244CC");
        Assert.True(double.Parse(right.Attribute("y2")!.Value, CultureInfo.InvariantCulture)
            - double.Parse(right.Attribute("y1")!.Value, CultureInfo.InvariantCulture) >= 39D);
        Assert.True(double.Parse(bottom.Attribute("x2")!.Value, CultureInfo.InvariantCulture)
            - double.Parse(bottom.Attribute("x1")!.Value, CultureInfo.InvariantCulture) >= 110D);
        var png = sheet.Range("A1:B2").ExportImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ShowGridlines = false });
        Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? pixels));
        Assert.NotNull(pixels);
        int longestRedColumn = Enumerable.Range(0, pixels!.Width).Max(x => Enumerable.Range(0, pixels.Height)
            .Count(y => { OfficeColor color = pixels.GetPixel(x, y); return color.R > 160 && color.G < 50 && color.B < 60; }));
        Assert.True(longestRedColumn >= 30, "The right border must span both merged rows in the raster output.");

        var html = new HtmlParser().ParseDocument(workbook.ToHtml());
        var merged = html.QuerySelector("[data-officeimo-cell='A1']")!;
        string style = merged.GetAttribute("style")!;
        Assert.Equal("2", merged.GetAttribute("rowspan"));
        Assert.Equal("2", merged.GetAttribute("colspan"));
        Assert.Equal("42.50", merged.TextContent);
        Assert.Contains("border-right:2px solid #CC1122", style, StringComparison.Ordinal);
        Assert.Contains("border-bottom:3px solid #2244CC", style, StringComparison.Ordinal);
        Assert.Contains("font-weight:700", style, StringComparison.Ordinal);
        Assert.Contains("background-color:#EEDDCC", style, StringComparison.Ordinal);
    }

    [Fact]
    public void BorderStylesThatCannotBeRepresentedExactlyProduceFidelityDiagnostics() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Borders");
        sheet.Cell(1, 1).SetValue("Dash-dot").SetBorder(ExcelCellBorders.Top, ExcelBorderStyle.DashDot, OfficeColor.Black);
        sheet.Cell(2, 1).SetValue("Diagonal").SetDiagonalBorder(ExcelBorderStyle.Thin, OfficeColor.Black);
        sheet.Cell(1, 3).SetValue("Segmented perimeter");
        sheet.Range("C1:D2").Merge();
        sheet.Cell(2, 3).SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Thin, OfficeColor.Parse("CC1122"));
        sheet.Cell(2, 4).SetBorder(ExcelCellBorders.Bottom, ExcelBorderStyle.Thick, OfficeColor.Parse("2244CC"));

        workbook.ToHtml(null, out OfficeConversionReport report);

        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Approximation
            && diagnostic.Location == "excel:cell:Borders!A1:top");
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission
            && diagnostic.Location == "excel:cell:Borders!A2:diagonal");
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission
            && diagnostic.Location == "Borders!C1:D2:bottom");
    }
}
