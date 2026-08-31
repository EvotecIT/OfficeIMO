using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Archive_message_count_is_enforced_before_archive_info_materialization() {
        byte[] validMessageInfo = Message(VarintField(1, 10000), VarintField(3, 0));
        byte[] archiveInfo = Message(new[] {
            VarintField(1, 1), BytesField(2, validMessageInfo)
        }.Concat(Enumerable.Range(0, 64)
            .Select(_ => BytesField(2, Array.Empty<byte>()))).ToArray());
        byte[] stream = Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo);
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(stream)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
                new IWorkReadOptions { MaximumRecordCount = 1 }));

        Assert.Contains("record count", exception.Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pages_section_entries_are_bounded_before_section_table_materialization() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        byte[] sectionTable = Message(Enumerable.Range(0, 64)
            .Select(_ => BytesField(1, Message())).ToArray());
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)));

        IWorkPagesProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages).ReadPages();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_SECTION_UNSUPPORTED");
    }

    [Fact]
    public void Image_reader_uses_a_later_complete_packaged_rendition() {
        const ulong originalIdentifier = 100;
        const ulong fallbackIdentifier = 101;
        const string originalName = "missing-original.png";
        const string fallbackName = "fallback.png";
        byte[] image = Message(BytesField(1, Message()),
            BytesField(11, Message(VarintField(1, originalIdentifier))),
            BytesField(12, Message(VarintField(1, fallbackIdentifier))));
        byte[] metadata = Message(
            BytesField(4, Message(VarintField(1, originalIdentifier),
                StringField(3, originalName), StringField(4, originalName))),
            BytesField(4, Message(VarintField(1, fallbackIdentifier),
                StringField(3, fallbackName), StringField(4, fallbackName))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message()),
            ArchiveRecord(10, 3005, image),
            ArchiveRecord(50, 11006, metadata));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ($"Data/{fallbackName}", ValidPreviewPng()));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Pages);
        IWorkArchiveRecord imageRecord = Assert.Single(source.Records,
            record => record.MessageType == 3005);

        IWorkImageAsset? asset = IWorkDrawingReader.ReadImage(source, imageRecord,
            new IWorkProjectionBudget(source.Options), out bool complete);

        Assert.True(complete);
        Assert.NotNull(asset);
        Assert.Equal($"Data/{fallbackName}", asset.PackagePath);
    }
}
