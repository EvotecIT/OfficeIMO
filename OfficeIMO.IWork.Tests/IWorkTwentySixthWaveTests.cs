using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Unsupported_keynote_drawables_with_text_use_visual_fallback() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong chartId = 5;
        const ulong storageId = 6;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(7, chartId))),
            ArchiveRecord(chartId, 7000, Message(ReferenceField(2, storageId)),
                new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Chart title"))));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED"
            && diagnostic.RecordIdentifier == chartId);
    }

    [Fact]
    public void Pages_list_placeholder_levels_adopt_their_authored_numbering_kind() {
        using MemoryStream package = CreatePagesPackageWithDescendingListLevels();
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Body body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("The reconstructed DOCX has no body.");
        Paragraph[] paragraphs = body.Elements<Paragraph>()
            .Where(paragraph => paragraph.InnerText is "Deep" or "Middle").ToArray();
        Assert.Equal(new[] { "Deep", "Middle" }, paragraphs.Select(paragraph => paragraph.InnerText));
        int numberId = paragraphs[0].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value
            ?? throw new InvalidDataException("The reconstructed list has no numbering identifier.");
        Assert.Equal(numberId,
            paragraphs[1].ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value);
        Numbering numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new InvalidDataException("The reconstructed DOCX has no numbering definitions.");
        int abstractId = numbering.Elements<NumberingInstance>()
            .Single(instance => instance.NumberID?.Value == numberId)
            .AbstractNumId?.Val?.Value
            ?? throw new InvalidDataException("The reconstructed list has no abstract definition.");
        Level[] levels = numbering.Elements<AbstractNum>()
            .Single(item => item.AbstractNumberId?.Value == abstractId)
            .Elements<Level>().ToArray();

        Assert.Equal(NumberFormatValues.LowerLetter,
            levels.Single(level => level.LevelIndex?.Value == 1).NumberingFormat?.Val?.Value);
        Assert.Equal(NumberFormatValues.Decimal,
            levels.Single(level => level.LevelIndex?.Value == 2).NumberingFormat?.Val?.Value);
    }

    [Fact]
    public void Pages_authored_oversized_image_extent_is_preserved() {
        using MemoryStream package = CreatePagesPackageWithImage(
            masked: false, left: 0, rotation: 0, width: 600, height: 900);

        using var result = WordDocument.LoadPagesWithReport(package);
        WordImage image = Assert.Single(result.Document.Images);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(600d, image.Width);
        Assert.Equal(900d, image.Height);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(600d, Assert.Single(reopened.Images).Width);
        Assert.Equal(900d, Assert.Single(reopened.Images).Height);
    }

    [Fact]
    public void Pages_image_extent_outside_the_docx_range_uses_visual_fallback() {
        using MemoryStream package = CreatePagesPackageWithImage(
            masked: false, left: 0, rotation: 0, width: float.MaxValue, height: 30);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    public void Zero_dimension_keynote_tables_use_visual_fallback(int rows, int columns) {
        using MemoryStream package = CreateKeynotePackageWithTableDefaults(
            rows, columns, defaultRowHeight: 10d, defaultColumnWidth: 30d,
            includePreview: true);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 0)]
    public void Zero_dimension_pages_tables_use_visual_fallback(int rows, int columns) {
        using MemoryStream package = CreatePagesPackageWithTableGeometry(
            0, 0, 0, 0, 0, includePreview: true, rows: rows, columns: columns);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    private static MemoryStream CreatePagesPackageWithDescendingListLevels() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong listStyleId = 3;
        const ulong deepStyleId = 4;
        const ulong middleStyleId = 5;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listStyleId))));
        byte[] paragraphTable = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, deepStyleId))),
            BytesField(1, Message(VarintField(1, 5), ReferenceField(2, middleStyleId))));
        byte[] listStyle = Message(
            VarintField(11, 1), VarintField(11, 1), VarintField(11, 1),
            FloatField(13, 0f), FloatField(13, 18f), FloatField(13, 36f),
            StringField(16, "•"), StringField(16, "a."), StringField(16, "1."));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Deep\nMiddle"), BytesField(5, paragraphTable),
                    BytesField(7, listTable)),
                new[] { listStyleId, deepStyleId, middleStyleId }),
            ArchiveRecord(listStyleId, 2023, listStyle),
            ArchiveRecord(deepStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, 36f))))),
            ArchiveRecord(middleStyleId, 2022,
                Message(BytesField(12, Message(FloatField(11, 18f))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
