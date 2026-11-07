using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 14D)]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas, 28D)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 14D)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable, 28D)]
    public void SaveAsPdf_RightToLeftReversesPhysicalColumnsWithTheirWidths(ExcelPdfWorksheetLayoutMode layout, double secondWidth) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("ColumnA");
        sheet.Cell(1, 2).SetValue("ColumnB");
        sheet.Cell(1, 1).SetFillColor(OfficeColor.Parse("FF0000"));
        sheet.Cell(1, 2).SetFillColor(OfficeColor.Parse("0000FF"));
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(14D));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(secondWidth));
        var options = CreateDirectionPdfOptions(layout);
        using PdfPigDocument leftToRight = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        sheet.View.RightToLeft = true;
        using PdfPigDocument rightToLeft = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));

        Assert.Equal(1, rightToLeft.NumberOfPages);
        var before = leftToRight.GetPage(1);
        var after = rightToLeft.GetPage(1);
        Assert.True(FindWordStartX(before, "ColumnA") < FindWordStartX(before, "ColumnB"));
        Assert.True(FindWordStartX(after, "ColumnB") < FindWordStartX(after, "ColumnA"));
        var beforeA = FindDirectionFillBounds(before, (1D, 0D, 0D));
        var beforeB = FindDirectionFillBounds(before, (0D, 0D, 1D));
        var afterA = FindDirectionFillBounds(after, (1D, 0D, 0D));
        var afterB = FindDirectionFillBounds(after, (0D, 0D, 1D));
        Assert.Equal(beforeA.Right - beforeA.Left, afterA.Right - afterA.Left, precision: 3);
        Assert.Equal(beforeB.Right - beforeB.Left, afterB.Right - afterB.Left, precision: 3);
        Assert.InRange(afterA.Left - afterB.Right, -0.01D, 0.01D);
        if (layout == ExcelPdfWorksheetLayoutMode.WorksheetCanvas) Assert.Equal(680D, afterA.Right, precision: 3);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_RightToLeftKeepsMergeAnchorsAndRepeatedTitlesOnTheStartingEdge(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("Title");
        sheet.Cell(1, 2).SetValue("Hidden");
        sheet.Cell(1, 3).SetValue("Merged");
        sheet.Range("C1:D1").Merge();
        sheet.Cell(1, 5).SetValue("BodyEnd");
        for (int column = 1; column <= 5; column++) sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(14D));
        sheet.SetColumnHidden(2, true);
        sheet.View.RightToLeft = true;
        sheet.PageSetup.PrintArea = "A1:E1";
        sheet.PageSetup.SetPrintTitles(firstRow: null, lastRow: null, firstColumn: 1, lastColumn: 1);
        sheet.AddManualColumnPageBreak(4);
        using PdfPigDocument pdf = PdfPigDocument.Open(workbook.ToPdf(CreateDirectionPdfOptions(layout)).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain("Hidden", pdf.GetPage(1).Text);
        Assert.True(FindWordStartX(pdf.GetPage(1), "Merged") < FindWordStartX(pdf.GetPage(1), "Title"));
        Assert.True(FindWordStartX(pdf.GetPage(2), "BodyEnd") < FindWordStartX(pdf.GetPage(2), "Title"));
        if (layout == ExcelPdfWorksheetLayoutMode.WorksheetCanvas)
            Assert.Equal(FindWordStartX(pdf.GetPage(1), "Title"), FindWordStartX(pdf.GetPage(2), "Title"), precision: 3);
    }

    [Fact]
    public void SaveAsPdf_RightToLeftMirrorsImagePlacementWithoutChangingItsSize() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("ImageColumn");
        sheet.Cell(1, 2).SetValue("OtherColumn");
        sheet.Cell(1, 1).SetFillColor(OfficeColor.Parse("FF0000"));
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(18D));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(28D));
        sheet.PageSetup.PrintArea = "A1:B3";
        sheet.AddImage(2, 1, CreateMinimalRgbPng(), new ExcelImageOptions {
            ContentType = "image/png", Width = OfficeLength.Points(18), Height = OfficeLength.Points(12),
            OffsetX = OfficeLength.Points(3)
        });
        var options = CreateDirectionPdfOptions(ExcelPdfWorksheetLayoutMode.WorksheetCanvas);
        byte[] beforeBytes = workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        sheet.View.RightToLeft = true;
        byte[] afterBytes = workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true });
        using PdfPigDocument before = PdfPigDocument.Open(beforeBytes);
        using PdfPigDocument after = PdfPigDocument.Open(afterBytes);
        var imageBefore = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(beforeBytes));
        var imageAfter = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(afterBytes));
        var cellBefore = FindDirectionFillBounds(before.GetPage(1), (1D, 0D, 0D));
        var cellAfter = FindDirectionFillBounds(after.GetPage(1), (1D, 0D, 0D));
        Assert.Equal(cellBefore.Left + 3D, imageBefore.X, precision: 3);
        Assert.Equal(cellAfter.Right - 3D, imageAfter.X + imageAfter.Width, precision: 3);
        Assert.Equal(18D, imageAfter.Width, precision: 3);
        Assert.Equal(12D, imageAfter.Height, precision: 3);
    }

    [Fact]
    public void SaveAsPdf_RightToLeftOversizedImageKeepsTheSameContinuationPageCount() {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("ColumnA");
        sheet.Cell(1, 2).SetValue("ColumnB");
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(14D));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(14D));
        sheet.AddImage(1, 1, CreateMinimalRgbPng(), new ExcelImageOptions {
            ContentType = "image/png", Width = OfficeLength.Points(1000), Height = OfficeLength.Points(20)
        });
        var options = CreateDirectionPdfOptions(ExcelPdfWorksheetLayoutMode.WorksheetCanvas);
        using PdfPigDocument leftToRight = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        sheet.View.RightToLeft = true;
        using PdfPigDocument rightToLeft = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));

        Assert.Equal(2, leftToRight.NumberOfPages);
        Assert.Equal(2, rightToLeft.NumberOfPages);
        Assert.Contains("ColumnA", rightToLeft.GetPage(1).Text);
        Assert.Contains("ColumnB", rightToLeft.GetPage(1).Text);
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_RightToLeftBoxedTableStyleKeepsVerticalBordersOnItsOutline(ExcelPdfWorksheetLayoutMode layout) {
        using ExcelDocument workbook = ExcelDocument.Create();
        ExcelSheet sheet = workbook.AddSheet("Report");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(1, 2).SetValue("Value");
        sheet.Cell(2, 1).SetValue("Item");
        sheet.Cell(2, 2).SetValue(1);
        sheet.SetColumnWidth(1, ExcelColumnWidth.Characters(14D));
        sheet.SetColumnWidth(2, ExcelColumnWidth.Characters(14D));
        sheet.AddTable("A1:B2", hasHeader: true, name: "Values", style: ExcelTableStyle.TableStyleLight9);
        ExcelPdfExportOptions options = CreateDirectionPdfOptions(layout);
        // Flow tables can also have a report-wide grid. Isolate the worksheet
        // table style's own outline in both directions.
        options.PdfOptions = new PdfCore.PdfCreateOptions {
            DefaultTableStyle = new PdfCore.PdfTableStyle { BorderColor = null }
        };
        using PdfPigDocument leftToRight = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        sheet.View.RightToLeft = true;
        using PdfPigDocument rightToLeft = PdfPigDocument.Open(workbook.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        foreach (PdfPigDocument pdf in new[] { leftToRight, rightToLeft }) {
            Assert.Equal(1, pdf.NumberOfPages);
            var boxes = pdf.GetPage(1).Paths.Select(path => path.GetBoundingRectangle())
                .Where(box => box.HasValue).Select(box => box!.Value).ToArray();
            var horizontal = boxes.Where(box => box.Width > 5D && box.Height < 1.5D).ToArray();
            var vertical = boxes.Where(box => box.Width < 1.5D && box.Height > 5D).ToArray();
            Assert.NotEmpty(horizontal);
            double left = horizontal.Min(box => box.Left);
            double right = horizontal.Max(box => box.Right);
            Assert.Contains(vertical, box => Math.Abs(box.Left - left) < 0.25D);
            Assert.Contains(vertical, box => Math.Abs(box.Right - right) < 0.25D);
            Assert.DoesNotContain(vertical, box => box.Left > left + 1D && box.Right < right - 1D);
        }
    }

    [Theory]
    [InlineData(ExcelPdfWorksheetLayoutMode.WorksheetCanvas)]
    [InlineData(ExcelPdfWorksheetLayoutMode.FlowTable)]
    public void SaveAsPdf_BoundedReadRetainsStoredRightToLeftDirection(ExcelPdfWorksheetLayoutMode layout) {
        byte[] bytes;
        using (ExcelDocument workbook = ExcelDocument.Create()) {
            ExcelSheet sheet = workbook.AddSheet("Report");
            sheet.Cell(1, 1).SetValue("ColumnA");
            sheet.Cell(1, 2).SetValue("ColumnB");
            sheet.View.RightToLeft = true;
            bytes = workbook.ToBytes();
        }
        using var input = new MemoryStream(bytes, writable: false);
        using ExcelDocument loaded = ExcelDocument.Load(input, new ExcelLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
        var worksheetPart = Assert.Single(loaded.OpenXmlDocument.WorkbookPart!.WorksheetParts);
        Assert.False(worksheetPart.IsRootElementLoaded);
        ExcelPdfExportOptions options = CreateDirectionPdfOptions(layout);
        options.UseBoundedWorksheetRead = true;
        options.MaxRowsPerSheet = 1;
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdf(options).ToBytes(new PdfCore.PdfSaveOptions { AllowLoss = true }));
        Assert.False(worksheetPart.IsRootElementLoaded);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.True(FindWordStartX(pdf.GetPage(1), "ColumnB") < FindWordStartX(pdf.GetPage(1), "ColumnA"));
        Assert.True(input.CanRead);
    }

    private static ExcelPdfExportOptions CreateDirectionPdfOptions(ExcelPdfWorksheetLayoutMode layout) => new() {
        IncludeSheetHeadings = false, HeaderRowCount = 0, WorksheetLayout = layout,
        PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(700), OfficeLength.Points(500)),
        Margins = PdfCore.PdfPageMargins.Uniform(OfficeLength.Points(20))
    };

    private static (double Left, double Right) FindDirectionFillBounds(UglyToad.PdfPig.Content.Page page, (double R, double G, double B) color) {
        var bounds = page.Paths.Where(path => path.IsFilled && path.FillColor.ToRGBValues() == color)
            .Select(path => path.GetBoundingRectangle()).Where(rectangle => rectangle.HasValue)
            .Select(rectangle => rectangle!.Value).ToArray();
        Assert.NotEmpty(bounds);
        return (bounds.Min(rectangle => rectangle.Left), bounds.Max(rectangle => rectangle.Right));
    }
}
