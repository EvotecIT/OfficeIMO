using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutExcelTests {
    [Fact]
    public void RegionsFromTableLimitedSheetsDoNotMoveOntoNarrativeSheets() {
        const string html = "<p>Narrative retained</p>" +
            "<table><caption>First</caption><tr><td>First table</td></tr></table>" +
            "<table><caption>Second</caption><tr><td>Second table" +
            "<div style='display:grid;width:140px;height:40px'>Limited layout</div></td></tr></table>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxTables = 1;
        limits.MaxSemanticContainers = 2;

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.DoesNotContain(workbook.Sheets, sheet => ContainsCellText(sheet, "Limited layout"));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("owning worksheet was not created", StringComparison.Ordinal));
    }

    [Fact]
    public void TableUnownedRegionsReceiveANarrativeWorksheet() {
        const string html = "<table><caption>Data</caption><tr><td>Table value</td></tr></table>" +
            "<div style='position:absolute;width:140px;height:40px'>Narrative layout</div>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.Equal("Imported", workbook.Sheets[1].Name);
        Assert.True(ContainsCellText(workbook.Sheets[1], "Narrative layout"));
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("owning worksheet was not created", StringComparison.Ordinal));
    }

    [Fact]
    public void RegionsInsideLaterRootTablesUseTheirOwningWorksheets() {
        const string html = "<table><caption>First</caption><tr><td>One<div style='display:grid;width:140px;height:40px'>First layout</div></td></tr></table>" +
            "<table><caption>Second</caption><tr><td>Two<div style='display:grid;width:140px;height:40px'>Second layout</div></td></tr></table>";

        HtmlToExcelResult result = HtmlConversionDocument.Parse(html)
            .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic });
        using ExcelDocument workbook = result.Value;

        Assert.Equal(2, workbook.Sheets.Count);
        Assert.True(ContainsCellText(workbook.Sheets[0], "First layout"));
        Assert.False(ContainsCellText(workbook.Sheets[0], "Second layout"));
        Assert.True(ContainsCellText(workbook.Sheets[1], "Second layout"));
    }

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
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;left:32px;top:200px;width:240px;height:72px'>First region" +
            "<img alt='First picture' src='" + image + "' style='width:12px;height:12px'></div>" +
            "<div style='position:absolute;left:32px;top:200px;width:240px;height:72px'>Second region" +
            "<img alt='Second picture' src='" + image + "' style='width:12px;height:12px'></div>";
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
        ExcelImage[] images = sheet.Images.ToArray();
        Assert.Equal(2, images.Length);
        ExcelImage firstImage = images[0];
        ExcelImage secondImage = images[1];
        Assert.True(firstImage.TryGetAbsoluteAnchorBounds(out _, out int firstY, out _, out _));
        Assert.True(secondImage.TryGetAbsoluteAnchorBounds(out _, out int secondY, out _, out _));
        Assert.True(secondY >= firstY + 80D);
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

    private static bool ContainsCellText(ExcelSheet sheet, string expected) {
        for (int row = 1; row <= 30; row++) {
            for (int column = 1; column <= 10; column++) {
                if (sheet.TryGetCellText(row, column, out string value)
                    && value.Contains(expected, StringComparison.Ordinal)) return true;
            }
        }
        return false;
    }
}
