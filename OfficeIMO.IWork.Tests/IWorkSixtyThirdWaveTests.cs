using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Duplicate_numbers_cell_offsets_disable_editable_reconstruction(
        bool wideOffsets) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Duplicate offsets", 1, 2, 42d,
                wideOffsets: wideOffsets, duplicatePopulatedOffset: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Empty(Assert.Single(Assert.Single(result.Projection.Sheets).Tables).Cells);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_text_formula_caches_round_trip_as_formula_strings() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Text formula", 1, 1, 0d, hasFormula: true,
                textValue: "Approved", completeFormula: true)
        });

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);
        ExcelSheet sheet = Assert.Single(result.Document.Sheets);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("1", sheet.GetFormulaText(1, 1));
        Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? cached));
        Assert.Equal("Approved", cached);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        ExcelSheet persisted = Assert.Single(reopened.Sheets);
        Assert.Equal("1", persisted.GetFormulaText(1, 1));
        Assert.True(persisted.TryGetCachedFormulaValue(1, 1, out string? persistedCache));
        Assert.Equal("Approved", persistedCache);
    }

    [Fact]
    public void Numbers_non_native_error_formula_caches_round_trip_as_formula_strings() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Error formula", 1, 1, 0d, hasFormula: true,
                error: true, completeFormula: true)
        });

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);
        ExcelSheet sheet = Assert.Single(result.Document.Sheets);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("1", sheet.GetFormulaText(1, 1));
        Assert.True(sheet.TryGetCachedFormulaValue(1, 1, out string? cached));
        Assert.Equal("#ERROR", cached);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        ExcelSheet persisted = Assert.Single(reopened.Sheets);
        Assert.Equal("1", persisted.GetFormulaText(1, 1));
        Assert.True(persisted.TryGetCachedFormulaValue(1, 1, out string? persistedCache));
        Assert.Equal("#ERROR", persistedCache);
    }

    [Theory]
    [InlineData(true, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    public void Pdf_dictionary_objects_reject_values_after_the_dictionary(
        bool trailCatalogDictionary, bool trailPagesDictionary,
        bool trailPageDictionary) {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            trailCatalogDictionary: trailCatalogDictionary,
            trailPagesDictionary: trailPagesDictionary,
            trailPageDictionary: trailPageDictionary);

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Fact]
    public void Pdf_dictionary_objects_allow_comments_before_the_terminator() {
        Assert.True(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true,
            commentCatalogDictionary: true)));
    }
}
