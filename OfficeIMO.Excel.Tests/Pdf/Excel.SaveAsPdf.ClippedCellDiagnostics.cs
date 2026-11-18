using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    private const string LongCellText = "Quarterly sales against target and plan for the region";

    /// <summary>Column A rows 1-6 and row 8 are blocked by column B; rows 10-12 are blocked in both A and B by column C.</summary>
    private static ExcelDocument CreateClippedCellWorkbook() {
        ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        foreach (int row in new[] { 1, 2, 3, 4, 5, 6, 8 }) {
            sheet.Cell(row, 1).SetValue(LongCellText);
            sheet.Cell(row, 2).SetValue("x");
        }
        for (int row = 10; row <= 12; row++) {
            sheet.Cell(row, 1).SetValue(LongCellText);
            sheet.Cell(row, 2).SetValue(LongCellText);
            sheet.Cell(row, 3).SetValue("x");
        }
        return document;
    }

    private static List<OfficeDiagnostic> CollectContentClipping(ExcelDocument document, ExcelPdfExportOptions? options = null) {
        using var stream = new MemoryStream();
        OfficeConversionReport report = document.SaveAsPdf(stream, options);
        Assert.True(report.HasLoss);
        return report.Diagnostics.Where(diagnostic => diagnostic.Code == "table-cell-content-clipped").ToList();
    }

    [Fact]
    public void SaveAsPdf_ReportsClippedCellsOncePerContiguousRange() {
        using ExcelDocument document = CreateClippedCellWorkbook();

        List<OfficeDiagnostic> clipped = CollectContentClipping(document);

        Assert.Equal(new[] { "Data!A1:A6", "Data!A8", "Data!A10:B12" }, clipped.Select(diagnostic => diagnostic.Location));
        Assert.All(clipped, diagnostic => Assert.Equal(OfficeConversionLossKind.Approximation, diagnostic.LossKind));
        Assert.StartsWith("6 cells (A1, A2, A3, ...). First cell: ", clipped.Single(diagnostic => diagnostic.Location == "Data!A1:A6").Message);
        Assert.StartsWith("1 cell (A8). First cell: ", clipped.Single(diagnostic => diagnostic.Location == "Data!A8").Message);
        Assert.StartsWith("6 cells (A10, B10, A11, ...). First cell: ", clipped.Single(diagnostic => diagnostic.Location == "Data!A10:B12").Message);
    }

    [Fact]
    public void SaveAsPdf_PerCellDiagnosticsReportEveryClippedCell() {
        using ExcelDocument document = CreateClippedCellWorkbook();

        List<OfficeDiagnostic> clipped = CollectContentClipping(document,
            new ExcelPdfExportOptions { DiagnosticsDetail = ExcelPdfDiagnosticsDetail.PerCell });

        Assert.Equal(13, clipped.Count);
        Assert.All(clipped, diagnostic => Assert.StartsWith("PdfTableCell[", diagnostic.Location));
        Assert.Equal(13, clipped.Select(diagnostic => diagnostic.Location).Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void SaveAsPdf_SummarizedClippingCoversTheSameCellsAsPerCellDiagnostics() {
        using ExcelDocument document = CreateClippedCellWorkbook();
        string[] codes = { "table-cell-content-clipped", "table-cell-paint-clipped", "table-cell-canvas-clip-paint-clipped" };
        using var summaryStream = new MemoryStream();
        using var perCellStream = new MemoryStream();

        OfficeConversionReport summary = document.SaveAsPdf(summaryStream);
        OfficeConversionReport perCell = document.SaveAsPdf(perCellStream,
            new ExcelPdfExportOptions { DiagnosticsDetail = ExcelPdfDiagnosticsDetail.PerCell });

        Assert.Equal(perCellStream.ToArray(), summaryStream.ToArray());
        foreach (string code in codes) {
            int perCellCount = perCell.Diagnostics.Count(diagnostic => diagnostic.Code == code);
            int summarizedCount = summary.Diagnostics.Where(diagnostic => diagnostic.Code == code)
                .Sum(diagnostic => int.Parse(diagnostic.Message.Substring(0, diagnostic.Message.IndexOf(' ')), System.Globalization.CultureInfo.InvariantCulture));
            Assert.Equal(perCellCount, summarizedCount);
        }
        Assert.Contains(perCell.Diagnostics, diagnostic => diagnostic.Code == "table-cell-paint-clipped");
    }
}
