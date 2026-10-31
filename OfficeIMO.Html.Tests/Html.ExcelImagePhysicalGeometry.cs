using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests;

public partial class HtmlOfficeAdapters {
    [Fact]
    public void ExcelHtml_InvalidNativePositionUsesItsPixelFallback() {
        using var workbook = ExcelDocument.Create();
        var sheet = workbook.AddSheet("Fallback image");
        sheet.CellValue(1, 1, "Seed");
        sheet.AddImage(OfficeLength.Pixels(12), OfficeLength.Pixels(7), OnePixelPng);
        string html = workbook.ToHtml(new ExcelHtmlSaveOptions { Profile = OfficeHtmlConversionProfile.ExcelSemanticTables });
        html = html.Replace("data-officeimo-x-emu=\"114300\"", "data-officeimo-x-emu=\"9223372036854775807\"")
            .Replace("data-officeimo-y-emu=\"66675\"", "data-officeimo-y-emu=\"-9223372036854775808\"");
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using var imported = result.Value;
        var image = Assert.Single(Assert.Single(imported.Sheets).Images);
        Assert.True(image.TryGetAbsoluteAnchorBounds(out var x, out var y, out _, out _));
        Assert.Equal(OfficeLength.Pixels(12), x);
        Assert.Equal(OfficeLength.Pixels(7), y);
        Assert.Equal(2, result.Report.Diagnostics.Count(value => value.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid));
    }

    [Fact]
    public void ExcelHtml_RoundTripsSignedNativeAbsoluteImagePosition() {
        using var source = new MemoryStream();
        using (var workbook = ExcelDocument.Create()) {
            var sheet = workbook.AddSheet("Absolute image");
            sheet.CellValue(1, 1, "Seed");
            sheet.AddImage(OfficeLength.Zero, OfficeLength.Zero, OnePixelPng, new OfficeIMO.Excel.ExcelImageOptions {
                Width = OfficeLength.Emus(9526),
                Height = OfficeLength.Emus(19051)
            });
            workbook.Save(source);
        }
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var position = drawing.Elements<Xdr.AbsoluteAnchor>().Single().Position!;
            position.X = -1L;
            position.Y = -2L;
            drawing.Save();
        }
        source.Position = 0;
        using var loaded = ExcelDocument.Load(source);
        string html = loaded.ToHtml(new ExcelHtmlSaveOptions { Profile = OfficeHtmlConversionProfile.ExcelSemanticTables });
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using var imported = result.Value;
        var image = Assert.Single(Assert.Single(imported.Sheets).Images);
        Assert.Empty(result.Report.Diagnostics);
        Assert.True(image.TryGetAbsoluteAnchorBounds(out var x, out var y, out var width, out var height));
        Assert.Equal(OfficeLength.Emus(-1), x);
        Assert.Equal(OfficeLength.Emus(-2), y);
        Assert.Equal(OfficeLength.Emus(9526), width);
        Assert.Equal(OfficeLength.Emus(19051), height);
    }

    [Theory]
    [InlineData(1L, 2L)]
    [InlineData(-1L, -2L)]
    public void ExcelHtml_RoundTripsNativeSameCellImageOffsets(long offsetX, long offsetY) {
        using var source = new MemoryStream();
        using (var workbook = ExcelDocument.Create()) {
            var sheet = workbook.AddSheet("Native markers");
            sheet.CellValue(1, 1, "Seed");
            sheet.AddImage("B2:B2", OnePixelPng, new OfficeIMO.Excel.ExcelImageOptions { OffsetX = OfficeLength.Pixels(-2), OffsetY = OfficeLength.Pixels(-3) })
                .SetTwoCellEndingMarker(2, 2);
            workbook.Save(source);
        }
        source.Position = 0;
        using (var package = SpreadsheetDocument.Open(source, true)) {
            var drawing = package.WorkbookPart!.WorksheetParts.Single().DrawingsPart!.WorksheetDrawing!;
            var marker = drawing.Elements<Xdr.TwoCellAnchor>().Single().ToMarker!;
            marker.ColumnOffset = new Xdr.ColumnOffset(offsetX.ToString(System.Globalization.CultureInfo.InvariantCulture));
            marker.RowOffset = new Xdr.RowOffset(offsetY.ToString(System.Globalization.CultureInfo.InvariantCulture));
            drawing.Save();
        }
        source.Position = 0;
        using var loaded = ExcelDocument.Load(source);
        string html = loaded.ToHtml(new ExcelHtmlSaveOptions { Profile = OfficeHtmlConversionProfile.ExcelSemanticTables });
        var result = HtmlConversionDocument.Parse(html).ToExcelDocumentResult();
        using var imported = result.Value;
        var image = Assert.Single(Assert.Single(imported.Sheets).Images);
        Assert.Empty(result.Report.Diagnostics);
        Assert.Equal(2, image.ToRowIndex);
        Assert.Equal(2, image.ToColumnIndex);
        Assert.Equal(OfficeLength.Emus(offsetX), image.ToOffsetX);
        Assert.Equal(OfficeLength.Emus(offsetY), image.ToOffsetY);
        using var saved = new MemoryStream();
        imported.Save(saved);
        saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        var savedImage = Assert.Single(Assert.Single(reopened.Sheets).Images);
        Assert.Equal(image.ToOffsetX, savedImage.ToOffsetX);
        Assert.Equal(image.ToOffsetY, savedImage.ToOffsetY);
    }
}
