using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Complete_uncached_numbers_formulas_do_not_fabricate_cached_values() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Uncached formula", 1, 1, 0d, hasFormula: true,
                formulaWithoutCachedValue: true, completeFormula: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        IWorkTableCell projectedCell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);
        ExcelCellData ownerCell = result.Document.Sheets[0].CellAt(1, 1).GetValue();

        Assert.False(result.IsVisualFallback);
        Assert.True(projectedCell.FormulaIsComplete);
        Assert.Null(projectedCell.Value);
        Assert.Equal(ExcelCellDataKind.Formula, ownerCell.Kind);
        Assert.NotNull(ownerCell.Formula);
        Assert.Null(ownerCell.Value);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        ExcelCellData persisted = reopened.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(ExcelCellDataKind.Formula, persisted.Kind);
        Assert.NotNull(persisted.Formula);
        Assert.Null(persisted.Value);
    }

    [Fact]
    public void Conflicting_pages_shape_text_storages_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "First",
            includePreview: true, alternateTextBox: "Second");

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Repeated_pages_shape_text_storage_fields_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Repeated",
            includePreview: true, duplicateTextBoxStorageField: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_shape_text_storage_aliases_across_fields_remain_editable() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Aliased",
            includePreview: true, aliasTextBoxStorageFields: true);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Contains(result.Document.Paragraphs, paragraph => paragraph.Text == "Aliased");
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Repeated_numbers_text_storage_references_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(Array.Empty<TableSpec>(), textBox: "Repeated",
            includePreview: true, duplicateTextBoxStorageReference: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_NUMBERS_TEXT_STORAGE_UNSUPPORTED");
    }
}
