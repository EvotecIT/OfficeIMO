using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_header_footer_variants_map_to_distinct_word_parts() {
        using MemoryStream package = CreatePagesPackageWithHeaderFooterVariants();

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        IWorkPagesSection source = Assert.Single(result.Projection.Sections);
        WordSection section = Assert.Single(result.Value.Sections);

        Assert.False(result.IsVisualFallback);
        Assert.True(source.HasFirstPageTemplate);
        Assert.True(source.HasEvenPageTemplate);
        Assert.True(source.HasDefaultPageTemplate);
        Assert.Equal("First header", Assert.Single(source.FirstPageHeaderContents).PlainText);
        Assert.Equal("Even header", Assert.Single(source.EvenPageHeaderContents).PlainText);
        Assert.Equal("Default header", Assert.Single(source.DefaultPageHeaderContents).PlainText);
        Assert.Contains(section.Header.First!.Paragraphs, paragraph => paragraph.Text == "First header");
        Assert.Contains(section.Header.Even!.Paragraphs, paragraph => paragraph.Text == "Even header");
        Assert.Contains(section.Header.Default!.Paragraphs, paragraph => paragraph.Text == "Default header");
        Assert.Contains(section.Footer.First!.Paragraphs, paragraph => paragraph.Text == "First footer");
        Assert.Contains(section.Footer.Even!.Paragraphs, paragraph => paragraph.Text == "Even footer");
        Assert.Contains(section.Footer.Default!.Paragraphs, paragraph => paragraph.Text == "Default footer");
        Assert.True(section.DifferentFirstPage);
        Assert.True(section.DifferentOddAndEvenPages);

        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        WordSection persisted = Assert.Single(reopened.Sections);
        Assert.Contains(persisted.Header.First!.Paragraphs, paragraph => paragraph.Text == "First header");
        Assert.Contains(persisted.Header.Even!.Paragraphs, paragraph => paragraph.Text == "Even header");
        Assert.Contains(persisted.Header.Default!.Paragraphs, paragraph => paragraph.Text == "Default header");
    }

    [Theory]
    [InlineData("1", PowerPointNumberingScheme.ArabicPlain, 1)]
    [InlineData("1.", PowerPointNumberingScheme.ArabicPeriod, 1)]
    [InlineData("a.", PowerPointNumberingScheme.AlphaLowerCharacterPeriod, 1)]
    [InlineData("iv.", PowerPointNumberingScheme.RomanLowerCharacterPeriod, 4)]
    public void Supported_keynote_ordered_markers_use_native_powerpoint_numbering(
        string label, PowerPointNumberingScheme expectedScheme, int expectedStart) {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Item", listLabel: label);

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        PowerPointParagraph paragraph = Assert.Single(Assert.Single(
            Assert.Single(result.Value.Slides).TextBoxes).Paragraphs);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("Item", paragraph.Text);
        Assert.True(paragraph.IsNumbered);
        Assert.Equal(expectedScheme, paragraph.NumberingScheme);
        Assert.Equal(expectedStart, paragraph.NumberingStartAt);
    }

    [Fact]
    public void Consecutive_keynote_list_items_continue_native_numbering() {
        using MemoryStream package = CreateKeynotePackageWithNumberedSequence("10.");

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        PowerPointParagraph[] paragraphs = Assert.Single(
            Assert.Single(result.Value.Slides).TextBoxes).Paragraphs.ToArray();

        Assert.Equal(2, paragraphs.Length);
        Assert.All(paragraphs, paragraph =>
            Assert.Equal(PowerPointNumberingScheme.ArabicPeriod, paragraph.NumberingScheme));
        Assert.Equal(10, paragraphs[0].NumberingStartAt);
        Assert.Null(paragraphs[1].NumberingStartAt);
        Assert.Equal(new[] { "One", "Two" }, paragraphs.Select(paragraph => paragraph.Text));
    }

    [Fact]
    public void Unsupported_keynote_ordered_markers_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: "Item", listLabel: "custom:");

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    private static MemoryStream CreatePagesPackageWithHeaderFooterVariants() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong firstTemplateId = 10;
        const ulong evenTemplateId = 11;
        const ulong defaultTemplateId = 12;
        byte[] sectionTable = Message(BytesField(1, Message(ReferenceField(2, sectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable)), new[] { sectionId }),
            ArchiveRecord(sectionId, 10011,
                Message(ReferenceField(23, firstTemplateId), ReferenceField(24, evenTemplateId),
                    ReferenceField(25, defaultTemplateId)),
                new[] { firstTemplateId, evenTemplateId, defaultTemplateId }),
            HeaderFooterTemplate(firstTemplateId, 20, 21),
            HeaderFooterTemplate(evenTemplateId, 22, 23),
            HeaderFooterTemplate(defaultTemplateId, 24, 25),
            ArchiveRecord(20, 2001, Message(StringField(3, "First header"))),
            ArchiveRecord(21, 2001, Message(StringField(3, "First footer"))),
            ArchiveRecord(22, 2001, Message(StringField(3, "Even header"))),
            ArchiveRecord(23, 2001, Message(StringField(3, "Even footer"))),
            ArchiveRecord(24, 2001, Message(StringField(3, "Default header"))),
            ArchiveRecord(25, 2001, Message(StringField(3, "Default footer"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));

        static byte[] HeaderFooterTemplate(ulong identifier, ulong headerId, ulong footerId) =>
            ArchiveRecord(identifier, 10143,
                Message(ReferenceField(1, headerId), ReferenceField(2, footerId)),
                new[] { headerId, footerId });
    }

    private static MemoryStream CreateKeynotePackageWithNumberedSequence(string label) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong listStyleId = 7;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "One\nTwo"), BytesField(7, listTable)),
                new[] { listStyleId }),
            ArchiveRecord(listStyleId, 2023,
                Message(VarintField(11, 1), StringField(16, label))));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)));
    }
}
