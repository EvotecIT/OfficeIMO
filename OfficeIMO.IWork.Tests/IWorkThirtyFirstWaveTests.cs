using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Keynote_first_observation_of_each_list_level_preserves_its_start(
        bool presenterNotes) {
        using MemoryStream package = CreateKeynotePackageWithNestedList(presenterNotes);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        PowerPointSlide slide = Assert.Single(result.Value.Slides);
        PowerPointParagraph[] paragraphs = (presenterNotes
                ? slide.Notes.Paragraphs
                : Assert.Single(slide.TextBoxes).Paragraphs)
            .Where(paragraph => paragraph.Text is "One" or "Deep")
            .ToArray();

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, paragraphs.Length);
        Assert.Equal(1, paragraphs[0].NumberingStartAt);
        Assert.Equal(3, paragraphs[1].NumberingStartAt);
        Assert.Equal(1, paragraphs[1].Level);

        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        PowerPointSlide persistedSlide = Assert.Single(reopened.Slides);
        PowerPointParagraph[] persisted = (presenterNotes
                ? persistedSlide.Notes.Paragraphs
                : Assert.Single(persistedSlide.TextBoxes).Paragraphs)
            .Where(paragraph => paragraph.Text is "One" or "Deep")
            .ToArray();
        Assert.Equal(2, persisted.Length);
        Assert.Equal(3, persisted[1].NumberingStartAt);
        Assert.Equal(1, persisted[1].Level);
    }

    [Theory]
    [InlineData('\u0004')]
    [InlineData('\u0005')]
    [InlineData('\u000c')]
    public void Pages_header_container_breaks_use_visual_fallback(char separator) {
        using MemoryStream package = CreatePagesPackageWithHeaderBreak(separator);

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
        Assert.Single(result.Value.Images);
    }

    [Fact]
    public void Pages_text_box_container_breaks_use_visual_fallback() {
        using MemoryStream package = CreatePagesPackage(includeBody: true,
            textBox: "Before\u000cAfter", includePreview: true);

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Fact]
    public void Keynote_slide_text_container_breaks_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Before\u000cAfter");

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }

    [Fact]
    public void Raster_preview_budget_is_spent_in_preference_order() {
        byte[] records = Message(ArchiveRecord(1, 10000, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("Preview-micro.png", CreateSizedPreviewPng(80, 100)),
            ("Preview.png", CreateSizedPreviewPng(100, 100)));
        var options = new IWorkReadOptions { MaximumPackageBytes = 41_000 };

        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages, options);

        Assert.Equal("Preview.png", source.PreferredRasterPreview!.Path);
        Assert.DoesNotContain(source.Previews,
            preview => preview.Path.Equals("Preview-micro.png", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Invalid_case_variant_does_not_hide_a_valid_preview() {
        byte[] records = Message(ArchiveRecord(1, 10000, Message()));
        byte[] malformed = CreateCrcValidPngWithInvalidImageData()
            .Concat(new byte[100]).ToArray();
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", malformed),
            ("Preview.PNG", ValidPreviewPng()));

        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages);

        Assert.Equal("Preview.PNG", Assert.Single(source.Previews).Path);
    }

    private static MemoryStream CreatePagesPackageWithHeaderBreak(char separator) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong templateId = 4;
        const ulong headerId = 5;
        byte[] sectionTable = Message(BytesField(1,
            Message(ReferenceField(2, sectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable)),
                new[] { sectionId }),
            ArchiveRecord(sectionId, 10011,
                Message(ReferenceField(25, templateId)), new[] { templateId }),
            ArchiveRecord(templateId, 10143,
                Message(ReferenceField(1, headerId)), new[] { headerId }),
            ArchiveRecord(headerId, 2001,
                Message(StringField(3, "Before" + separator + "After"))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithNestedList(bool presenterNotes) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeOrNoteId = 5;
        const ulong storageId = 6;
        const ulong listStyleId = 7;
        const ulong firstParagraphStyleId = 8;
        const ulong secondParagraphStyleId = 9;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] paragraphTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, firstParagraphStyleId))),
            BytesField(1, Message(VarintField(1, 4), ReferenceField(2, secondParagraphStyleId))));
        byte[] listStyle = Message(
            VarintField(11, 1), VarintField(11, 1),
            FloatField(13, 0f), FloatField(13, 18f),
            StringField(16, "1."), StringField(16, "c."));
        var records = new List<byte[]> {
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(2, showId)), new[] { showId }),
            ArchiveRecord(showId, 2, KeynoteShow(slideTree), new[] { nodeId }),
            ArchiveRecord(nodeId, 4,
                Message(ReferenceField(2, slideId)), new[] { slideId }),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(presenterNotes ? 27 : 5, shapeOrNoteId)),
                new[] { shapeOrNoteId }),
            ArchiveRecord(shapeOrNoteId, presenterNotes ? 15u : 2011u,
                Message(ReferenceField(presenterNotes ? 1 : 2, storageId)),
                new[] { storageId }),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "One\nDeep"), BytesField(5, paragraphTable),
                    BytesField(7, listTable)),
                new[] { listStyleId, firstParagraphStyleId, secondParagraphStyleId }),
            ArchiveRecord(listStyleId, 2023, listStyle),
            ArchiveRecord(firstParagraphStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, 0f))))),
            ArchiveRecord(secondParagraphStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, 18f)))))
        };
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }
}
