using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void WorksheetPrintCentering_MovesPdfAndPngContentInsideMargins(bool horizontal, bool vertical) {
        string path = Path.Combine(_directoryWithFiles, $"CenteredLayout-{horizontal}-{vertical}.xlsx");
        using var workbook = ExcelDocument.Create(path, new ExcelCreateOptions { FirstSheetName = "Centered" });
        ExcelSheet sheet = workbook.Sheets[0];
        sheet.CellValue(1, 1, "Marker");
        sheet.CellValue(2, 2, "Tail");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(12));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(12));
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
        sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(30)));
        sheet.SetMargins(OfficeLength.Inches(0.5), OfficeLength.Inches(1), OfficeLength.Inches(0.75), OfficeLength.Inches(1.25));
        sheet.SetPageSetup(paperSize: ExcelPaperSize.Letter, scale: 50);
        sheet.Range("A1:B2").SetFillColor("FFCC33");
        sheet.SetPrintOptions(horizontalCentered: false, verticalCentered: false);
        byte[] originalPdf = workbook.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var original = PdfPigDocument.Open(originalPdf);
        var pngOriginal = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, SplitByManualPageBreaks = true })[0];
        sheet.SetPrintOptions(horizontalCentered: horizontal, verticalCentered: vertical);
        byte[] centeredPdf = workbook.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var centered = PdfPigDocument.Open(centeredPdf);
        var pngCentered = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, SplitByManualPageBreaks = true })[0];
        var before = original.GetPage(1);
        var after = centered.GetPage(1);
        // Measure the uncentered cell fills independently of the centering code.
        // Character-based column widths depend on the available workbook font.
        var fills = before.Paths
            .Where(path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 0.8D, 0.2D))
            .Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue)
            .Select(bounds => bounds!.Value)
            .ToArray();
        Assert.NotEmpty(fills);
        double width = fills.Max(bounds => bounds.Right) - fills.Min(bounds => bounds.Left);
        double height = fills.Max(bounds => bounds.Top) - fills.Min(bounds => bounds.Bottom);
        Assert.Equal(30D, height, precision: 3);
        double expectedX = horizontal ? (612 - 36 - 72 - width) / 2 : 0;
        double expectedY = vertical ? (792 - 54 - 90 - height) / 2 : 0;
        Assert.InRange(FindWordStartX(after, "Marker") - FindWordStartX(before, "Marker"), expectedX - 0.05, expectedX + 0.05);
        Assert.InRange(FindWordStartY(before, "Marker") - FindWordStartY(after, "Marker"), expectedY - 0.05, expectedY + 0.05);
        Assert.True(OfficeRasterImageDecoder.TryDecode(pngOriginal.Bytes, out var imageBefore));
        Assert.True(OfficeRasterImageDecoder.TryDecode(pngCentered.Bytes, out var imageAfter));
        var boxBefore = FindYellowPrintBounds(imageBefore!);
        var boxAfter = FindYellowPrintBounds(imageAfter!);
        double pngExpectedX = horizontal ? (pngCentered.Width - 144 - boxBefore.Width) / 2D : 0;
        double pngExpectedY = vertical ? (pngCentered.Height - 192 - boxBefore.Height) / 2D : 0;
        Assert.InRange(boxAfter.X - boxBefore.X, pngExpectedX - 2, pngExpectedX + 2);
        Assert.InRange(boxAfter.Y - boxBefore.Y, pngExpectedY - 2, pngExpectedY + 2);
        string? evidenceDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_EXCEL_INTEGRATION_EVIDENCE_DIR");
        if (!string.IsNullOrWhiteSpace(evidenceDirectory)) {
            Directory.CreateDirectory(evidenceDirectory);
            string name = $"centering-{horizontal}-{vertical}";
            File.WriteAllBytes(Path.Combine(evidenceDirectory, name + "-before.pdf"), originalPdf);
            File.WriteAllBytes(Path.Combine(evidenceDirectory, name + "-after.pdf"), centeredPdf);
            File.WriteAllBytes(Path.Combine(evidenceDirectory, name + "-before.png"), pngOriginal.Bytes);
            File.WriteAllBytes(Path.Combine(evidenceDirectory, name + "-after.png"), pngCentered.Bytes);
        }
    }

    private static (int X, int Y, int Width, int Height) FindYellowPrintBounds(OfficeRasterImage image) {
        int left = image.Width, top = image.Height, right = -1, bottom = -1;
        for (int y = 0; y < image.Height; y++) {
            for (int x = 0; x < image.Width; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.R < 240 || pixel.G < 180 || pixel.G > 220 || pixel.B > 70) continue;
                left = Math.Min(left, x); top = Math.Min(top, y);
                right = Math.Max(right, x); bottom = Math.Max(bottom, y);
            }
        }
        Assert.True(right >= left && bottom >= top, "Worksheet fill was not painted.");
        return (left, top, right - left + 1, bottom - top + 1);
    }
}
