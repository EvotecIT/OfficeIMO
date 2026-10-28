using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public void SaveAsPdf_ManualBreaksSeparateChartsOutsideThePopulatedCells(bool columnBreak, int cellPosition) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Charts");
        var data = new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Values", new[] { 1D, 2D }) });
        sheet.AddChart(data, 2, 2, 240, 180, title: "FirstChart");
        sheet.AddChart(data, columnBreak ? 2 : 22, columnBreak ? 22 : 2, 240, 180, title: "SecondChart");
        if (cellPosition == 1) sheet.CellValue(1, 1, "EarlyCell");
        if (cellPosition == 2) sheet.CellValue(columnBreak ? 2 : 31, columnBreak ? 31 : 2, "LateCell");
        sheet.SetPageSetup(fitToWidth: columnBreak ? 0U : 1U, fitToHeight: columnBreak ? 1U : 0U);
        if (columnBreak) sheet.AddManualColumnPageBreak(21);
        else sheet.AddManualRowPageBreak(21);
        using var package = new MemoryStream(workbook.ToBytes());
        using ExcelDocument loaded = ExcelDocument.Load(package);
        Assert.Equal(new[] { 21 }, columnBreak
            ? loaded.Sheets[0].GetManualColumnPageBreaks() : loaded.Sheets[0].GetManualRowPageBreaks());
        using PdfCore.PdfDocument converted = loaded.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PdfPageSize(2000, 800), Margins = PdfCore.PdfPageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(converted.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FirstChart", pdf.GetPage(1).Text);
        Assert.DoesNotContain("SecondChart", pdf.GetPage(1).Text);
        Assert.Contains("SecondChart", pdf.GetPage(2).Text);
        Assert.DoesNotContain("FirstChart", pdf.GetPage(2).Text);
        if (cellPosition == 1) Assert.Contains("EarlyCell", pdf.GetPage(1).Text);
        if (cellPosition == 2) Assert.Contains("LateCell", pdf.GetPage(2).Text);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_CompactLeadingChartStaysWhole(bool fit) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Compact");
        sheet.CellValue(5, 1, "Anchor");
        sheet.AddChart(new ExcelChartData(new[] { "A", "B" },
            new[] { new ExcelChartSeries("Values", new[] { 1D, 2D }) }), 2, 2, 240, 180, title: "CompactChart");
        sheet.SetPageSetup(fitToWidth: fit ? 1U : 0U, fitToHeight: fit ? 1U : 0U);
        using PdfCore.PdfDocument converted = workbook.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PdfPageSize(2000, 800), Margins = PdfCore.PdfPageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(converted.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("CompactChart", pdf.GetPage(1).Text);
        Assert.Contains("Anchor", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(ExcelPageOrder.OverThenDown, false)]
    [InlineData(ExcelPageOrder.DownThenOver, false)]
    [InlineData(ExcelPageOrder.OverThenDown, true)]
    [InlineData(ExcelPageOrder.DownThenOver, true)]
    public void SaveAsPdf_OrdersCellChunksAndMediaTilesTogether(ExcelPageOrder order, bool titles) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddWorksheet("Order");
        sheet.CellValue(1, 1, "LeftCell"); sheet.CellValue(1, 2, "RightCell");
        sheet.SetColumnWidth(1, 70); sheet.SetColumnWidth(2, 70);
        var data = new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Values", new[] { 1D, 2D }) });
        sheet.AddChart(data, 2, 1, 200, 180, title: "TopLeft");
        sheet.AddChart(data, 2, 2, 200, 180, title: "TopRight");
        sheet.AddChart(data, 22, 1, 200, 180, title: "BottomLeft");
        sheet.AddChart(data, 22, 2, 200, 180, title: "BottomRight");
        if (titles) workbook.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1);
        sheet.AddManualColumnPageBreak(1); sheet.AddManualRowPageBreak(21);
        sheet.SetPageSetup(fitToWidth: 0U, fitToHeight: 0U, pageOrder: order);
        using PdfCore.PdfDocument converted = workbook.ToPdf(new ExcelPdfExportOptions {
            IncludeSheetHeadings = false, HeaderRowCount = 0,
            PageSize = new PdfCore.PdfPageSize(titles ? 2000 : 500, 800), Margins = PdfCore.PdfPageMargins.Uniform(20)
        });
        using PdfPigDocument pdf = PdfPigDocument.Open(converted.ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        string[] expected = order == ExcelPageOrder.OverThenDown
            ? new[] { "TopLeft", "TopRight", "BottomLeft", "BottomRight" }
            : new[] { "TopLeft", "BottomLeft", "TopRight", "BottomRight" };
        Assert.Equal(expected.Length, pdf.NumberOfPages);
        for (int index = 0; index < expected.Length; index++) {
            Assert.Contains(expected[index], pdf.GetPage(index + 1).Text);
            foreach (string other in expected.Where(title => title != expected[index])) Assert.DoesNotContain(other, pdf.GetPage(index + 1).Text);
        }
    }

}
