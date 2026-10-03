using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Shared_invalid_presenter_note_storages_are_omitted_once_across_slides() {
        const int slideCount = 128;
        const int storageCount = 64;
        var records = new List<byte[]>();
        var slideNodes = new List<byte[]>();
        for (int index = 0; index < slideCount; index++) {
            ulong nodeId = (ulong)(1000 + index * 2);
            ulong slideId = nodeId + 1;
            slideNodes.Add(ReferenceField(2, nodeId));
            records.Add(ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))));
            records.Add(ArchiveRecord(slideId, 5, Message(ReferenceField(27, 10))));
        }
        records.Insert(0, ArchiveRecord(1, 1, Message(ReferenceField(2, 2))));
        records.Insert(1, ArchiveRecord(2, 2, KeynoteShow(Message(slideNodes.ToArray()))));
        records.Add(ArchiveRecord(10, 15, Message(Enumerable.Range(0, storageCount)
            .Select(index => ReferenceField(1, (ulong)(100 + index))).ToArray())));
        for (int index = 0; index < storageCount; index++)
            records.Add(ArchiveRecord((ulong)(100 + index), 2001, new byte[] { 0x80 }));

        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote).ReadKeynote();
        IWorkConversionReport report = projection.CreateConversionReport(
            IWorkProjectionKind.EditableReconstruction, null, Array.Empty<IWorkDiagnostic>(),
            allowPartialEditableReconstruction: true);

        Assert.Equal(slideCount, projection.Slides.Count);
        Assert.Equal(storageCount, report.SourceUnits.Count(unit =>
            unit.Kind == IWorkSourceUnitKind.Text && unit.Disposition == IWorkSourceUnitDisposition.Omitted));
    }

    [Fact]
    public void Distinct_invalid_presenter_note_storages_keep_one_diagnostic_each() {
        const int slideCount = 128;
        var records = new List<byte[]>();
        var slideNodes = new List<byte[]>();
        for (int index = 0; index < slideCount; index++) {
            ulong nodeId = (ulong)(1000 + index * 4);
            ulong slideId = nodeId + 1;
            ulong noteId = nodeId + 2;
            ulong storageId = nodeId + 3;
            slideNodes.Add(ReferenceField(2, nodeId));
            records.Add(ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))));
            records.Add(ArchiveRecord(slideId, 5, Message(ReferenceField(27, noteId))));
            records.Add(ArchiveRecord(noteId, 15, Message(ReferenceField(1, storageId))));
            records.Add(ArchiveRecord(storageId, 2001, new byte[] { 0x80 }));
        }
        records.Insert(0, ArchiveRecord(1, 1, Message(ReferenceField(2, 2))));
        records.Insert(1, ArchiveRecord(2, 2, KeynoteShow(Message(slideNodes.ToArray()))));

        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote).ReadKeynote();

        Assert.Equal(slideCount, projection.Slides.Count);
        Assert.Equal(slideCount, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED"));
        Assert.Equal(slideCount, projection.Diagnostics
            .Where(diagnostic => diagnostic.Code == "IWORK_KEYNOTE_TEXT_STORAGE_UNSUPPORTED")
            .Select(diagnostic => diagnostic.RecordIdentifier).Distinct().Count());
    }

    [Fact]
    public void Distinct_invalid_pages_shape_storages_keep_one_diagnostic_each() {
        const int shapeCount = 128;
        var records = new List<byte[]>();
        var documentReferences = new List<ulong> { 2 };
        for (int index = 0; index < shapeCount; index++) {
            ulong shapeId = (ulong)(1000 + index * 2);
            ulong storageId = shapeId + 1;
            documentReferences.Add(shapeId);
            records.Add(ArchiveRecord(shapeId, 2011, Message(
                BytesField(1, Message(BytesField(1, GeometryDrawable(72f, 72f, 120f, 40f)))),
                ReferenceField(2, storageId))));
            records.Add(ArchiveRecord(storageId, 2001, Message(BytesField(3, new byte[] { 0xc3, 0x28 }))));
        }
        records.Insert(0, ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), documentReferences));
        records.Insert(1, ArchiveRecord(2, 2001, Message(StringField(3, "Body"))));

        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))));
        IWorkPagesProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();

        Assert.Equal(shapeCount, projection.Diagnostics.Count(diagnostic =>
            diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED"));
        Assert.Equal(shapeCount, projection.Diagnostics
            .Where(diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED")
            .Select(diagnostic => diagnostic.RecordIdentifier).Distinct().Count());
    }
}
