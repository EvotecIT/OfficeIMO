using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Numbers_catalogs_reject_fields_outside_the_entry_envelope(bool formula) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Catalog", 1, 1, 1d,
                textValue: formula ? null : "Value",
                completeFormula: formula,
                unexpectedStringCatalogFieldCount: formula ? 0 : 1,
                unexpectedFormulaCatalogFieldCount: formula ? 1 : 0)
        });

        Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers)
                .ReadNumbers());
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

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

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

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

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

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
    }
}
