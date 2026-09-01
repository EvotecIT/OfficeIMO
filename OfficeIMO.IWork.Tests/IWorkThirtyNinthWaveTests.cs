using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_formula_catalog_is_independent_of_the_materialized_cell_limit() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula catalog", 1, 1, 42d,
                hasFormula: true, duplicateFormula: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        IWorkNumbersProjection projection = source.ReadNumbers();

        Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
    }

    [Fact]
    public void Numbers_string_catalog_is_independent_of_the_materialized_cell_limit() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("String catalog", 1, 1, 0d,
                textValue: "Value", duplicateString: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        IWorkNumbersProjection projection = source.ReadNumbers();

        Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
    }
}
