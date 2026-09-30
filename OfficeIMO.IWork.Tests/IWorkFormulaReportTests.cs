using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Formula_report_keeps_expression_and_cache_assessments_independent() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Same", 1, 1, 42d, hasFormula: true),
            new TableSpec("Same", 1, 1, 84d, hasFormula: true, completeFormula: true),
            new TableSpec("Ordinary", 1, 1, 21d)
        });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false },
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.Empty(result.Report.PreservedRecords);
        Assert.Equal(2, result.Report.FormulaSummary.TotalCount);
        Assert.Equal(1, result.Report.FormulaSummary.CompleteExpressionCount);
        Assert.Equal(1, result.Report.FormulaSummary.IncompleteExpressionCount);
        Assert.Equal(2, result.Report.FormulaSummary.CompleteCacheCount);
        Assert.Equal(0, result.Report.FormulaSummary.PartialCacheCount);
        Assert.Equal(0, result.Report.FormulaSummary.MissingCacheCount);
        Assert.Equal(new ulong[] { 10, 14 }, result.Report.FormulaCells.Select(cell => cell.TableIdentity!.RecordIdentifier));
        Assert.All(result.Report.FormulaCells, cell => {
            Assert.Equal(1, cell.Row); Assert.Equal(1, cell.Column);
            Assert.Equal(IWorkCellKind.Number, cell.CachedValueKind);
        });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(ExcelCellDataKind.Formula, reopened.Sheets[1].CellAt(1, 1).GetValue().Kind);
    }
}
