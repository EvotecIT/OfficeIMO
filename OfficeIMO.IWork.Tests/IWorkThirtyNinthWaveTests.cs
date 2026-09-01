using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Numbers_formula_catalog_is_bounded_before_entries_are_materialized() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Formula catalog", 1, 1, 42d,
                hasFormula: true, duplicateFormula: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("formula catalog", exception.Message, StringComparison.Ordinal);
        Assert.Contains("remaining table-catalog limit of 1", exception.Message,
            StringComparison.Ordinal);
    }

    [Fact]
    public void Numbers_string_catalog_is_bounded_before_entries_are_materialized() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("String catalog", 1, 1, 0d,
                textValue: "Value", duplicateString: true)
        });
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumMaterializedCells = 1 });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());

        Assert.Contains("string catalog", exception.Message, StringComparison.Ordinal);
        Assert.Contains("remaining table-catalog limit of 1", exception.Message,
            StringComparison.Ordinal);
    }
}
