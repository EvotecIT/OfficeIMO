using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(1_048_576, true)]
    [InlineData(1_048_577, false)]
    public void Numbers_text_box_projection_respects_the_xlsx_row_limit(
        int textBoxCount, bool expected) {
        Assert.Equal(expected, ExcelIWorkConverter.FitsTextBoxesInWorksheet(textBoxCount));
    }

    [Fact]
    public void Keynote_natural_alignment_uses_the_rtl_text_direction() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "مرحبا بالعالم", naturalAlignment: true);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);
        PowerPointParagraph paragraph = Assert.Single(Assert.Single(
            Assert.Single(result.Document.Slides).TextBoxes).Paragraphs);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(PowerPointTextAlignment.Right, paragraph.Alignment);
        Assert.True(paragraph.RightToLeft);
    }

    [Fact]
    public void Pages_natural_alignment_uses_the_rtl_text_direction() {
        using MemoryStream package = CreatePagesPackageWithStyleChain(
            depth: 1, naturalAlignment: true, bodyText: "مرحبا بالعالم");

        using var result = WordIWorkConverter.LoadPagesWithReport(package);
        WordParagraph paragraph = Assert.Single(result.Document.Paragraphs,
            candidate => candidate.Text == "مرحبا بالعالم");

        Assert.False(result.IsVisualFallback);
        Assert.Equal(WordParagraphAlignment.Start, paragraph.ParagraphAlignment);
        Assert.True(paragraph.BiDi);
    }

    [Theory]
    [InlineData("IIII.")]
    [InlineData("IIV.")]
    [InlineData("iV.")]
    public void Noncanonical_keynote_roman_markers_use_visual_fallback(string label) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Item", listLabel: label);

        using var result = PowerPointIWorkConverter.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData("IIII.")]
    [InlineData("IIV.")]
    [InlineData("iV.")]
    public void Noncanonical_pages_roman_markers_use_visual_fallback(string label) {
        using MemoryStream package = CreatePagesPackageWithListLabel(label,
            includePreview: true);

        using var result = WordIWorkConverter.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }
}
