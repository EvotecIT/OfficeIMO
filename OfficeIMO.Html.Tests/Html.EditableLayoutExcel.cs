using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutExcelTests {
    [Fact]
    public void PositionedAndGridRegionsReopenAsEditableCellAnchors() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:200px;width:240px;height:72px;background-color:#dbeafe;" +
            "background-image:url('" + image + "'),url('" + image + "');background-repeat:no-repeat;background-size:18px 18px;" +
            "background-position:left top,right bottom;box-shadow:2px 2px 4px #555}" +
            ".grid{display:grid;grid-template-columns:1fr 1fr;width:300px;height:80px;background:#fef3c7}" +
            "</style><div class='positioned'>Editable positioned<img src='" + image + "' style='opacity:.4;width:24px;height:18px'></div>" +
            "<div class='grid'><span>Grid A</span><span>Grid B</span></div>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using ExcelDocument reopened = ExcelDocument.Load(
            new MemoryStream(stream.ToArray()),
            new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        ExcelSheet sheet = Assert.Single(reopened.Sheets);

        Assert.True(sheet.TryGetCellText(13, 2, out string positioned));
        Assert.Equal("Editable positioned", positioned);
        Assert.Equal("DBEAFE", sheet.CellAt(13, 2).GetStyle().FillColorHex);
        Assert.True(sheet.TryGetCellText(3, 1, out string grid));
        Assert.Equal("Grid AGrid B", grid);
        Assert.Equal("FEF3C7", sheet.CellAt(3, 1).GetStyle().FillColorHex);
        Assert.Contains(sheet.GetMergedRanges(), range => range.A1Range == "B13:E16");
        Assert.Contains(sheet.GetMergedRanges(), range => range.A1Range == "A3:E6");
        Assert.Contains(sheet.Images, drawing => drawing.TransparencyPercent == 60);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void OverlappingRegionsMoveToDistinctNonOverlappingCellAnchors() {
        const string html = "<div style='position:absolute;left:32px;top:200px;width:240px;height:72px'>First region</div>"
            + "<div style='position:absolute;left:32px;top:200px;width:240px;height:72px'>Second region</div>";
        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;
        ExcelSheet sheet = Assert.Single(workbook.Sheets);

        Assert.True(sheet.TryGetCellText(13, 2, out string first));
        Assert.True(sheet.TryGetCellText(18, 2, out string second));
        Assert.Equal("First region", first);
        Assert.Equal("Second region", second);
        Assert.Contains(sheet.GetMergedRanges(), range => range.A1Range == "B13:E16");
        Assert.Contains(sheet.GetMergedRanges(), range => range.A1Range == "B18:E21");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void LayoutPicturesRespectSharedImageAndShapeBudgets() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;top:120px;width:120px;height:60px'>Budgeted region"
            + "<img src='" + image + "' style='width:12px;height:12px'><img src='" + image
            + "' style='width:12px;height:12px'></div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxShapes = 1;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Single(Assert.Single(workbook.Sheets).Images);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }
}
