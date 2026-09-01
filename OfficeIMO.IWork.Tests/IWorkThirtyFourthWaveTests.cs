using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Record_limit_is_enforced_before_nested_message_infos_are_parsed() {
        byte[] validMessageInfo = Message(VarintField(1, 10000), VarintField(3, 0));
        byte[] archiveInfo = Message(
            VarintField(1, 1),
            BytesField(2, validMessageInfo),
            BytesField(2, new byte[] { 0x80 }));
        byte[] stream = Message(Varint(checked((ulong)archiveInfo.Length)), archiveInfo);
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(stream)));

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
                new IWorkReadOptions { MaximumRecordCount = 1 }));

        Assert.Contains("record count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Pages_text_storage_shared_by_a_header_and_drawable_is_projected_twice() {
        using MemoryStream package = CreatePagesPackageWithSharedHeaderDrawableStorage();

        IWorkPagesProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages).ReadPages();

        Assert.True(projection.HasEditableContent);
        Assert.Equal("Shared text", Assert.Single(Assert.Single(
            projection.Sections).DefaultPageHeaderContents).PlainText);
        Assert.Equal("Shared text", Assert.Single(projection.TextBoxes));
    }

    [Fact]
    public void Numbers_cells_cannot_consume_bytes_from_the_next_populated_offset() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Crossing", 1, 2, 42d, cellCrossesNextOffset: true)
        }, includePreview: true);

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_DECODE");
    }

    [Fact]
    public void Keynote_rejects_wrong_type_presenter_note_wrappers() {
        using MemoryStream package = CreateKeynotePackageWithWrongTypePresenterNote();

        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Keynote).ReadKeynote();

        Assert.False(projection.HasEditableContent);
        Assert.Empty(Assert.Single(projection.Slides).PresenterNotes);
        Assert.Contains(projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_NOTES_UNSUPPORTED");
    }

    private static MemoryStream CreatePagesPackageWithSharedHeaderDrawableStorage() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong sectionId = 3;
        const ulong headerFooterId = 4;
        const ulong storageId = 5;
        const ulong zOrderId = 6;
        const ulong shapeId = 7;
        byte[] sectionTable = Message(BytesField(1,
            Message(ReferenceField(2, sectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId), ReferenceField(20, zOrderId)),
                new[] { bodyId, zOrderId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable)),
                new[] { sectionId }),
            ArchiveRecord(sectionId, 10011,
                Message(ReferenceField(25, headerFooterId)), new[] { headerFooterId }),
            ArchiveRecord(headerFooterId, 10143,
                Message(ReferenceField(1, storageId)), new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Shared text"))),
            ArchiveRecord(zOrderId, 9000,
                Message(ReferenceField(1, shapeId)), new[] { shapeId }),
            ArchiveRecord(shapeId, 2011,
                Message(BytesField(1, Message(BytesField(1,
                            GeometryDrawable(10f, 10f, 100f, 50f)))),
                    ReferenceField(2, storageId)), new[] { storageId }));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateKeynotePackageWithWrongTypePresenterNote() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong noteId = 5;
        const ulong storageId = 6;
        byte[] slideTree = Message(ReferenceField(2, nodeId));
        byte[] records = Message(
            ArchiveRecord(documentId, 1,
                Message(ReferenceField(2, showId)), new[] { showId }),
            ArchiveRecord(showId, 2,
                KeynoteShow(slideTree), new[] { nodeId }),
            ArchiveRecord(nodeId, 4,
                Message(ReferenceField(2, slideId)), new[] { slideId }),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(27, noteId)), new[] { noteId }),
            ArchiveRecord(noteId, 2011,
                Message(ReferenceField(1, storageId)), new[] { storageId }),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Not a note"))));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)));
    }
}
