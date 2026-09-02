using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_even_header_mode_populates_every_section() {
        using MemoryStream package = CreatePagesPackageWithSectionWideEvenHeaders();

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Value.Sections.Count);
        Assert.Contains(result.Value.Sections[0].Header.Even!.Paragraphs,
            paragraph => paragraph.Text == "First default header");
        Assert.Contains(result.Value.Sections[0].Footer.Even!.Paragraphs,
            paragraph => paragraph.Text == "First default footer");
        Assert.Contains(result.Value.Sections[1].Header.Even!.Paragraphs,
            paragraph => paragraph.Text == "Second even header");
        Assert.Contains(result.Value.Sections[1].Footer.Even!.Paragraphs,
            paragraph => paragraph.Text == "Second even footer");

        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Contains(reopened.Sections[0].Header.Even!.Paragraphs,
            paragraph => paragraph.Text == "First default header");
        Assert.Contains(reopened.Sections[1].Header.Even!.Paragraphs,
            paragraph => paragraph.Text == "Second even header");
    }

    [Fact]
    public void Empty_Pages_text_shapes_retain_link_metadata() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: string.Empty,
            includePreview: true,
            textBoxDrawable: Message(StringField(4, "https://example.test/empty-pages-shape")));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Equal("https://example.test/empty-pages-shape",
            Assert.Single(result.Projection.TextBoxObjects).Hyperlink);
        Assert.Contains(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Empty_Keynote_text_shapes_retain_link_metadata() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            text: string.Empty,
            textBoxDrawable: Message(StringField(4, "https://example.test/empty-keynote-shape")));

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback);
        Assert.Equal("https://example.test/empty-keynote-shape",
            Assert.Single(result.Projection.Slides).TitleBox!.Hyperlink);
        Assert.Equal(new Uri("https://example.test/empty-keynote-shape"),
            Assert.Single(Assert.Single(result.Value.Slides).TextBoxes).Hyperlink);
    }

    private static MemoryStream CreatePagesPackageWithSectionWideEvenHeaders() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstSectionId = 3;
        const ulong secondSectionId = 4;
        const ulong firstDefaultTemplateId = 10;
        const ulong secondEvenTemplateId = 11;
        const ulong secondDefaultTemplateId = 12;
        byte[] sectionTable = Message(
            BytesField(1, Message(ReferenceField(2, firstSectionId))),
            BytesField(1, Message(ReferenceField(2, secondSectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "First\u0004Second"), BytesField(17, sectionTable)),
                new[] { firstSectionId, secondSectionId }),
            ArchiveRecord(firstSectionId, 10011,
                Message(ReferenceField(25, firstDefaultTemplateId)),
                new[] { firstDefaultTemplateId }),
            ArchiveRecord(secondSectionId, 10011,
                Message(ReferenceField(24, secondEvenTemplateId),
                    ReferenceField(25, secondDefaultTemplateId)),
                new[] { secondEvenTemplateId, secondDefaultTemplateId }),
            HeaderFooterTemplate(firstDefaultTemplateId, 20, 21),
            HeaderFooterTemplate(secondEvenTemplateId, 22, 23),
            HeaderFooterTemplate(secondDefaultTemplateId, 24, 25),
            ArchiveRecord(20, 2001, Message(StringField(3, "First default header"))),
            ArchiveRecord(21, 2001, Message(StringField(3, "First default footer"))),
            ArchiveRecord(22, 2001, Message(StringField(3, "Second even header"))),
            ArchiveRecord(23, 2001, Message(StringField(3, "Second even footer"))),
            ArchiveRecord(24, 2001, Message(StringField(3, "Second default header"))),
            ArchiveRecord(25, 2001, Message(StringField(3, "Second default footer"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));

        static byte[] HeaderFooterTemplate(ulong identifier, ulong headerId, ulong footerId) =>
            ArchiveRecord(identifier, 10143,
                Message(ReferenceField(1, headerId), ReferenceField(2, footerId)),
                new[] { headerId, footerId });
    }
}
