using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void SaveAsPdf_FutureTitleColumnsRetainBodyImagesAndCharts(int anchorColumn) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int column = 1; column <= 6; column++) {
            sheet.Cell(1, column, "Column" + column);
            sheet.SetColumnWidth(column, 10D);
        }
        sheet.AddImage(2, anchorColumn, CreateMinimalRgbPng(), "image/png", widthPixels: 24, heightPixels: 16);
        sheet.AddChart(new ExcelChartData(new[] { "One", "Two" }, new[] { new ExcelChartSeries("Values", new double[] { 1, 2 }) }),
            3, anchorColumn, widthPixels: 120, heightPixels: 80, title: "RetainedChart");
        document.SetPrintArea(sheet, "A1:F4");
        document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 3, lastCol: 3);
        sheet.AddManualColumnPageBreak(2);

        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(700, 500), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Column3", pdf.GetPage(1).Text);
        Assert.Contains("Column3", pdf.GetPage(2).Text);
        Assert.Contains("RetainedChart", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RetainedChart", pdf.GetPage(2).Text);
        PdfCore.PdfImagePlacement image = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(1, image.PageNumber);
        Assert.InRange(image.X, 20D, 100D);
        Assert.Equal(18D, image.Width, precision: 3);
    }

    [Fact]
    public void SaveAsPdf_TerminalMediaUsesDisplayedColumnsAndRealTrailingGap() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Report");
        for (int column = 1; column <= 15; column++) sheet.SetColumnWidth(column, 10D);
        sheet.Cell(1, 1, "Title"); sheet.Cell(1, 12, "FinalBody");
        // M lies directly after the body; O additionally includes the real empty N column.
        sheet.AddImage(2, 13, CreateMinimalRgbPng(), "image/png", widthPixels: 24, heightPixels: 16);
        sheet.AddImage(2, 15, CreateMinimalRgbPng(), "image/png", widthPixels: 24, heightPixels: 16);
        document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 1, lastCol: 1);
        sheet.AddManualColumnPageBreak(10);

        byte[] bytes = document.ToPdfBytes(new ExcelToPdfOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PageSize(800, 500), Margins = PdfCore.PageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        PdfCore.PdfImagePlacement[] images = PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes).OrderBy(image => image.X).ToArray();
        Assert.Equal(2, images.Length);
        Assert.All(images, image => Assert.Equal(2, image.PageNumber));
        Assert.InRange(images[0].X, 150D, 220D);
        Assert.InRange(images[1].X - images[0].X, 100D, 120D);
        Assert.All(images, image => Assert.Equal(18D, image.Width, precision: 3));
    }
}
