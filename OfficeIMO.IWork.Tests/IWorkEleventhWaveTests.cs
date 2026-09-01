using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Enforces_the_source_wide_projected_image_budget() {
        var budget = new IWorkProjectionBudget(new IWorkReadOptions { MaximumProjectedImages = 1 });

        budget.AddImage();
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => budget.AddImage());

        Assert.Contains("image count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Editable_owners_accept_only_the_image_formats_they_can_materialize() {
        Assert.True(IWorkDrawingReader.IsEditableOwnerImageMediaType("image/png"));
        Assert.True(IWorkDrawingReader.IsEditableOwnerImageMediaType("image/jpeg"));
        Assert.False(IWorkDrawingReader.IsEditableOwnerImageMediaType("image/svg+xml"));
        Assert.False(IWorkDrawingReader.IsEditableOwnerImageMediaType("application/pdf"));
    }

    [Fact]
    public void Jpeg_previews_require_a_complete_entropy_decode() {
        using FileStream input = File.OpenRead(Fixture("nim-iwork/simple.pages"));
        using var fixture = new System.IO.Compression.ZipArchive(input,
            System.IO.Compression.ZipArchiveMode.Read, leaveOpen: false);
        byte[] jpeg = ReadEntry(fixture, "preview.jpg");
        int scan = FindJpegMarker(jpeg, 0xda);
        int scanLength = jpeg[scan + 2] << 8 | jpeg[scan + 3];
        byte[] truncated = jpeg.Take(scan + 2 + scanLength)
            .Concat(new byte[] { 0xff, 0xd9 }).ToArray();

        (int? width, int? height) = IWorkImageInfo.Read(
            truncated, "image/jpeg", 64L * 1024 * 1024);

        Assert.Null(width);
        Assert.Null(height);
    }

    [Fact]
    public void Inline_object_markers_disable_editable_text_reconstruction() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, bodyText: "Before\ufffcAfter");

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    [Fact]
    public void Odd_length_numbers_cell_offsets_disable_editable_reconstruction() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Odd offsets", 1, 1, 42d, oddCurrentOffsets: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Keynote_owner_preserves_slide_names_native_numbering_and_inline_breaks() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "First\u2028Second", slideName: "Named slide", listLabel: "10.");

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        PowerPointSlide slide = Assert.Single(result.Value.Slides);
        PowerPointParagraph paragraph = Assert.Single(Assert.Single(slide.TextBoxes).Paragraphs);

        Assert.Equal("Named slide", slide.Name);
        Assert.Equal("First\nSecond", paragraph.Text);
        Assert.True(paragraph.IsNumbered);
        Assert.Equal(PowerPointNumberingScheme.ArabicPeriod, paragraph.NumberingScheme);
        Assert.Equal(10, paragraph.NumberingStartAt);
        Assert.Contains(paragraph.InlineNodes,
            node => node.Kind == PowerPointParagraphInlineKind.LineBreak);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        PowerPointSlide persisted = Assert.Single(reopened.Slides);
        Assert.Equal("Named slide", persisted.Name);
        PowerPointParagraph persistedParagraph = Assert.Single(
            Assert.Single(persisted.TextBoxes).Paragraphs);
        Assert.Equal(PowerPointNumberingScheme.ArabicPeriod, persistedParagraph.NumberingScheme);
        Assert.Equal(10, persistedParagraph.NumberingStartAt);
        Assert.Contains(persistedParagraph.InlineNodes,
            node => node.Kind == PowerPointParagraphInlineKind.LineBreak);
    }

    [Fact]
    public void Pages_text_box_accessibility_descriptions_round_trip_through_word() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: "Accessible box",
            includePreview: false, textBoxDrawable: Message(StringField(8, "Source description")));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        Assert.Equal("Source description", Assert.Single(result.Value.TextBoxes).Description);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal("Source description", Assert.Single(reopened.TextBoxes).Description);
    }

    private static int FindJpegMarker(byte[] bytes, byte marker) {
        for (int index = 0; index + 1 < bytes.Length; index++) {
            if (bytes[index] == 0xff && bytes[index + 1] == marker) return index;
        }
        throw new InvalidDataException($"JPEG marker 0x{marker:x2} was not found.");
    }
}
