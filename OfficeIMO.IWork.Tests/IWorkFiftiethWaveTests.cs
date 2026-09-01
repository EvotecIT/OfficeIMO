using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Malformed_numbers_catalogs_use_visual_fallback(bool formula, bool malformedWire) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Catalog", 1, 1, 1d,
                textValue: formula ? null : "Value",
                completeFormula: formula,
                unexpectedStringCatalogFieldCount: !formula && !malformedWire ? 1 : 0,
                unexpectedFormulaCatalogFieldCount: formula && !malformedWire ? 1 : 0,
                malformedStringCatalog: !formula && malformedWire,
                malformedFormulaCatalog: formula && malformedWire)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == (formula
                ? "IWORK_TABLE_FORMULA_STORAGE_UNSUPPORTED"
                : "IWORK_TABLE_STRING_STORAGE_UNSUPPORTED"));
    }

    [Fact]
    public void Keynote_slide_tree_rejects_fields_outside_the_reference_envelope() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            unexpectedSlideTreeFieldCount: 1);

        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote).ReadKeynote();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_TREE_MISSING");
    }

    [Fact]
    public void Mixed_case_keynote_alphabetic_markers_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Item", listLabel: "Ab.");

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Keynote_sub_hundredth_paragraph_spacing_uses_visual_fallback(bool before) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            spaceBefore: before ? 1.234f : null,
            spaceAfter: before ? null : 1.234f);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Keynote_hundredth_point_paragraph_spacing_remains_editable(bool before) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            spaceBefore: before ? 1.23f : null,
            spaceAfter: before ? null : 1.23f);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
    }
}
