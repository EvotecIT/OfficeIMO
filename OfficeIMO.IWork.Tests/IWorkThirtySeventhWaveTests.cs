using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(71f, 540f)]
    [InlineData(4033f, 540f)]
    [InlineData(960f, 71f)]
    [InlineData(960f, 4033f)]
    public void Keynote_slide_sizes_outside_the_presentationml_range_use_a_valid_default_canvas(
        float width, float height) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            slideWidth: width, slideHeight: height);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Equal(960d, result.Document.SlideSize.WidthPoints, 3);
        Assert.Equal(540d, result.Document.SlideSize.HeightPoints, 3);
        Assert.Empty(result.Document.ValidateDocument());
    }

    [Theory]
    [InlineData(72f)]
    [InlineData(4032f)]
    public void Keynote_slide_sizes_at_the_presentationml_boundaries_remain_editable(float size) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            slideWidth: size, slideHeight: size);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(size, result.Document.SlideSize.WidthPoints, 3);
        Assert.Equal(size, result.Document.SlideSize.HeightPoints, 3);
        Assert.Empty(result.Document.ValidateDocument());
    }

    [Fact]
    public void Numbers_rejects_excess_tile_rows_before_parsing_nested_row_messages() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Bounded rows", 1, 1, 1d, malformedSecondTileRow: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_ROW_COUNT_UNSUPPORTED");
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_ROWS_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_rejects_excess_tiles_before_parsing_nested_tile_entries() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Bounded tiles", 1, 1, 1d, malformedSecondTileEntry: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_TILE_COUNT_UNSUPPORTED");
        Assert.DoesNotContain(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Duplicate_keynote_drawables_within_one_field_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            duplicateDrawableInField: true);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DUPLICATE_DRAWABLE");
    }

    [Fact]
    public void Keynote_drawables_cannot_fill_both_title_and_body_roles() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            aliasDrawableAcrossFields: true);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }
}
