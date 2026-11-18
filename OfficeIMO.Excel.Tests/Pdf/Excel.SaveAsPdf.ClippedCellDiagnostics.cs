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

    [Fact]
    public void SaveAsPdf_MergedCellAcrossPageBreakReportsItsAnchorCell() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Report");
        for (int row = 1; row <= 4; row++) sheet.SetRowHeight(row, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(36)));
        sheet.Cell(1, 1).SetValue("Start");
        sheet.Cell(2, 1).SetValue(LongCellText);
        sheet.Range("A2:A3").Merge();
        sheet.Cell(2, 2).SetValue("x");
        sheet.Cell(3, 2).SetValue("x");
        sheet.Cell(4, 2).SetValue("End");
        sheet.PageSetup.PrintArea = "A1:B4";
        sheet.AddManualRowPageBreak(2);
        string[] codes = { "table-cell-content-clipped", "table-cell-paint-clipped", "table-cell-canvas-clip-paint-clipped" };
        using var summaryStream = new MemoryStream();
        using var perCellStream = new MemoryStream();

        OfficeConversionReport summary = document.SaveAsPdf(summaryStream);
        OfficeConversionReport perCell = document.SaveAsPdf(perCellStream,
            new ExcelPdfExportOptions { DiagnosticsDetail = ExcelPdfDiagnosticsDetail.PerCell });

        // Both page fragments of the merge report clipping; the summary reports the same codes, each once,
        // counting the merge as its one anchor cell.
        List<OfficeDiagnostic> occurrences = perCell.Diagnostics.Where(diagnostic => codes.Contains(diagnostic.Code)).ToList();
        Assert.Contains(occurrences, diagnostic => diagnostic.Location!.IndexOf(" on page 1 ", StringComparison.Ordinal) >= 0);
        Assert.Contains(occurrences, diagnostic => diagnostic.Location!.IndexOf(" on page 2 ", StringComparison.Ordinal) >= 0);
        List<OfficeDiagnostic> summarized = summary.Diagnostics.Where(diagnostic => codes.Contains(diagnostic.Code)).ToList();
        Assert.Equal(occurrences.Select(diagnostic => diagnostic.Code).Distinct(StringComparer.Ordinal).OrderBy(code => code, StringComparer.Ordinal),
            summarized.Select(diagnostic => diagnostic.Code).OrderBy(code => code, StringComparer.Ordinal));
        Assert.All(summarized, diagnostic => {
            Assert.Equal("Report!A2", diagnostic.Location);
            Assert.StartsWith("1 cell (A2). First cell: ", diagnostic.Message);
        });
    }

    [Fact]
    public void SaveAsPdf_FlowTableLayoutSummarizesClippedCells() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue(LongCellText + " " + LongCellText);
        sheet.Range("A1:A2").Merge();
        sheet.Cell(1, 2).SetValue("x");
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(10)));
        sheet.SetRowHeight(2, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(10)));
        using var summaryStream = new MemoryStream();
        using var perCellStream = new MemoryStream();

        OfficeConversionReport summary = document.SaveAsPdf(summaryStream,
            new ExcelPdfExportOptions { WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable });
        OfficeConversionReport perCell = document.SaveAsPdf(perCellStream, new ExcelPdfExportOptions {
            WorksheetLayout = ExcelPdfWorksheetLayoutMode.FlowTable, DiagnosticsDetail = ExcelPdfDiagnosticsDetail.PerCell });

        // The flow table paints the merge once per row it spans; per cell, each row is a finding.
        Assert.Equal(2, perCell.Diagnostics.Count(diagnostic => diagnostic.Code == "table-cell-paint-clipped"));
        OfficeDiagnostic clipped = Assert.Single(summary.Diagnostics, diagnostic => diagnostic.Code == "table-cell-paint-clipped");
        Assert.Equal("Data!A1", clipped.Location);
        Assert.StartsWith("1 cell (A1). First cell: ", clipped.Message);
    }
}
