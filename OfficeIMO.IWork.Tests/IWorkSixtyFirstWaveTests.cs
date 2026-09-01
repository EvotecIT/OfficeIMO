using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Covered_merge_detection_scales_with_maximum_sparse_input() {
        const int count = 100_000;
        IWorkTableMergeRange[] merges = Enumerable.Range(1, count)
            .Select(row => new IWorkTableMergeRange(row, 1, row, 2))
            .ToArray();
        IWorkTableCell[] cells = Enumerable.Range(1, count)
            .Select(row => new IWorkTableCell(row, 3, IWorkCellKind.Number, (double)row))
            .ToArray();
        var table = new IWorkTable("Sparse", count, 3, cells, mergedRanges: merges);

        Assert.False(table.HasPopulatedCoveredMergeCells());
    }

    [Fact]
    public void Numbers_dates_beyond_xlsx_tick_precision_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Date", 1, 1, 0.0000001d, date: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.IsType<DateTime>(Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables[0].Cells).Value);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_formula_dates_beyond_xlsx_tick_precision_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula date", 1, 1, 0.0000001d, hasFormula: true,
                date: true, completeFormula: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            result.Projection.Sheets).Tables[0].Cells);
        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal(IWorkCellKind.DateTime, cell.ValueKind);
    }

    [Fact]
    public void Numbers_dates_at_xlsx_precision_remain_editable() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Date", 1, 1, 0.001d, date: true)
        });

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.False(result.IsVisualFallback);
    }
}
