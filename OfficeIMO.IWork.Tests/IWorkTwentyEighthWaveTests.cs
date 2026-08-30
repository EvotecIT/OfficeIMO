using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Explicit_source_kind_overrides_a_misleading_path_extension() {
        string directory = Path.Combine(Path.GetTempPath(),
            "OfficeIMO-IWork-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "renamed.pages");
        try {
            using MemoryStream package = CreatePackage(("Index/Document.iwa",
                FrameIwa(Message(ArchiveRecord(1, 1, Message())))));
            File.WriteAllBytes(path, package.ToArray());

            IWorkSourceDocument source = IWorkSourceDocument.Open(
                path, IWorkDocumentKind.Numbers);

            Assert.Equal(IWorkDocumentKind.Numbers, source.Kind);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(5, 2022)]
    [InlineData(7, 2023)]
    [InlineData(8, 2021)]
    [InlineData(11, 2032)]
    public void Duplicate_rich_text_attribute_offsets_disable_editable_reconstruction(
        int tableField, int archiveType) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstAttributeId = 3;
        const ulong secondAttributeId = 4;
        byte[] table = Message(
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, firstAttributeId))),
            BytesField(1, Message(VarintField(1, 0), ReferenceField(2, secondAttributeId))));
        byte[] attribute = archiveType switch {
            2023 => Message(VarintField(11, 1)),
            2032 => Message(StringField(2, "https://example.test/")),
            _ => Message()
        };
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Text"), BytesField(tableField, table)),
                new[] { firstAttributeId, secondAttributeId }),
            ArchiveRecord(firstAttributeId, checked((uint)archiveType), attribute),
            ArchiveRecord(secondAttributeId, checked((uint)archiveType), attribute));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.False(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData(9)]
    [InlineData(10)]
    [InlineData(14)]
    public void Keynote_paragraph_pagination_flags_are_explicitly_diagnosed(int styleField) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong styleId = 7;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                Message(BytesField(3, Message(ReferenceField(2, nodeId))))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Paragraph"), BytesField(5, styleTable)),
                new[] { styleId }),
            ArchiveRecord(styleId, 2022,
                Message(BytesField(12, Message(VarintField(styleField, 1))))));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        Assert.Contains(result.ImportReport.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_UNSUPPORTED");
    }
}
