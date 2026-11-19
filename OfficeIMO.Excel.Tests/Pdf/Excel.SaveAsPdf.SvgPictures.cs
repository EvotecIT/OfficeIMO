using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void SaveAsPdf_DrawsSvgPictureAsVectorsInsteadOfItsRasterFallback() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<rect width='40' height='20' fill='#ff0000'/></svg>");
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= 6; row++) sheet.Cell(row, 1).SetValue(row);
        sheet.AddImage(2, 2, svg, new ExcelImageOptions { ContentType = "image/svg+xml",
            Width = global::OfficeIMO.OfficeLength.Points(80), Height = global::OfficeIMO.OfficeLength.Points(40) });

        byte[] pdf = document.ToPdf().ToBytes();

        Assert.Empty(PdfCore.PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(Regex.Matches(Encoding.GetEncoding("ISO-8859-1").GetString(pdf), @"/Subtype /Form /FormType 1 /BBox \[0 0 40 20\]"));
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_SvgPictureKeepsItsCropAndLinkAsVectorsInsideItsPlacement(ExcelPdfWorksheetLayoutMode layout) {
        // Red left half, blue right half; the crop keeps only the blue half.
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<rect width='20' height='20' fill='#ff0000'/><rect x='20' width='20' height='20' fill='#0000ff'/></svg>");
        using var document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        for (int row = 1; row <= 4; row++) sheet.Cell(row, 1).SetValue(row);
        ExcelImage image = sheet.AddImage(2, 2, svg, new ExcelImageOptions { ContentType = "image/svg+xml", Description = "Order code",
            Width = global::OfficeIMO.OfficeLength.Points(80), Height = global::OfficeIMO.OfficeLength.Points(40) });
        image.SetCropRatio(.5D, 0D, 0D, 0D);
        image.HyperlinkUri = new Uri("https://example.com/order/42");

        byte[] pdf = document.ToPdf(new ExcelPdfExportOptions { WorksheetLayout = layout, IncludeSheetHeadings = false,
            PageSize = new PdfCore.PdfPageSize(global::OfficeIMO.OfficeLength.Points(300), global::OfficeIMO.OfficeLength.Points(200)),
            Margins = PdfCore.PdfPageMargins.Uniform(global::OfficeIMO.OfficeLength.Points(20)) }).ToBytes();

        Assert.Empty(PdfCore.PdfDocument.Load(pdf).Reader.Images());
        PdfCore.PdfLogicalLinkAnnotation link = Assert.Single(PdfCore.PdfDocumentReadResult.Load(pdf).GetLinksByUri("https://example.com/order/42"));
        Assert.Equal("Order code", link.Contents);
        OfficeRasterImage page = OfficeDrawingRasterRenderer.Render(PdfCore.PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        int red = 0, left = int.MaxValue, top = int.MaxValue, right = -1, bottom = -1;
        for (int y = 0; y < page.Height; y++) {
            for (int x = 0; x < page.Width; x++) {
                OfficeColor color = page.GetPixel(x, y);
                if (color == OfficeColor.FromRgb(255, 0, 0)) red++;
                if (color != OfficeColor.FromRgb(0, 0, 255)) continue;
                left = Math.Min(left, x);
                top = Math.Min(top, y);
                right = Math.Max(right, x + 1);
                bottom = Math.Max(bottom, y + 1);
            }
        }

        Assert.Equal(0, red);
        // The cropped half fills the whole picture frame, which is also the link area.
        Assert.InRange(Math.Abs(left - link.SourceLink.X1), 0D, 1.5D);
        Assert.InRange(Math.Abs(right - link.SourceLink.X2), 0D, 1.5D);
        Assert.InRange(Math.Abs(top - (200D - link.SourceLink.Y2)), 0D, 1.5D);
        Assert.InRange(Math.Abs(bottom - (200D - link.SourceLink.Y1)), 0D, 1.5D);
        Assert.InRange(link.SourceLink.Width / link.SourceLink.Height, 1.95D, 2.05D);
    }

    [Fact]
    public void SaveAsPdf_SvgPictureThatCannotBeDrawnExportsItsRasterFallbackAndReportsIt() {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='20'>" +
            "<rect width='40' height='20' fill='#ff0000'/></svg>");
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            document.AddSheet("Data").AddImage(2, 2, svg, new ExcelImageOptions { ContentType = "image/svg+xml",
                Width = global::OfficeIMO.OfficeLength.Points(80), Height = global::OfficeIMO.OfficeLength.Points(40) });
            document.Save();
        }
        OfficeSvgPartDamage.Truncate(stream);

        using var reopened = ExcelDocument.Load(stream);
        byte[] pdf = reopened.ToPdf(null, out OfficeConversionReport report).ToBytes();

        Assert.Single(PdfCore.PdfDocument.Load(pdf).Reader.Images());
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.SvgImageRasterFallback);
    }
}
