using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_column_widths_preserve_converted_precision_in_the_excel_owner() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Width", 1, 1, 1d, defaultColumnWidth: 73d)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.False(result.IsVisualFallback);
        double width = Assert.IsType<double>(
            result.Document.Sheets[0].DefaultColumnWidth);
        Assert.Equal(13.19047619047619d, width, 12);
    }

    [Fact]
    public void Mixed_case_pages_alphabetic_markers_use_visual_fallback() {
        using MemoryStream package = CreatePagesPackageWithListLabel(
            "Ab.", includePreview: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Fact]
    public void Covered_merge_cells_report_materialized_content() {
        var merge = new IWorkTableMergeRange(1, 1, 1, 2);
        var anchorOnly = new IWorkTable("Anchor", 1, 2, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Number, 1d)
        }, mergedRanges: new[] { merge });
        var populatedCoveredCell = new IWorkTable("Covered", 1, 2, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Number, 1d),
            new IWorkTableCell(1, 2, IWorkCellKind.Number, 2d)
        }, mergedRanges: new[] { merge });

        Assert.False(anchorOnly.HasPopulatedCoveredMergeCells());
        Assert.True(populatedCoveredCell.HasPopulatedCoveredMergeCells());
    }
}
