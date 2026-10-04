using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Numbers_objects_selected_as_sheets_with_an_unsupported_type_retain_their_identity(bool visual) {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2), ReferenceField(1, 5))),
            ArchiveRecord(2, 2, Message(StringField(1, "Recovered"))),
            ArchiveRecord(5, 9999, Message()), ArchiveRecord(99, 9999, Message())))),
            ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual);
        IWorkSourceUnit unit = Assert.Single(report.SourceUnits,
            unit => unit.Kind == IWorkSourceUnitKind.UnsupportedObject);
        Assert.Equal(5ul, unit.Identity.RecordIdentifier);
        Assert.Equal(visual ? IWorkSourceUnitDisposition.Unassessed : IWorkSourceUnitDisposition.Omitted,
            unit.Disposition);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 99);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Selected_unsupported_objects_are_counted_once_and_inactive_records_are_excluded(
        IWorkDocumentKind kind, bool visual) {
        byte[] selected = Message(ReferenceField(1, 10), ReferenceField(1, 10), ReferenceField(1, 11));
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), ReferenceField(20, 3)), new ulong[] { 2, 3 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
                ArchiveRecord(3, 10020, selected, new ulong[] { 10, 11 })),
            IWorkDocumentKind.Numbers => Message(
                ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10),
                    ReferenceField(2, 10), ReferenceField(2, 11)))),
            _ => Message(
                ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(6, 10), ReferenceField(6, 10), ReferenceField(6, 11))))
        };
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, 9999, Message()), ArchiveRecord(11, 9998, Message()),
            ArchiveRecord(99, 9999, Message())))), ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual);
        IWorkSourceUnitCount count = Count(report, IWorkSourceUnitKind.UnsupportedObject);
        Assert.Equal(2, count.TotalCount);
        Assert.Equal(0, count.ReconstructedCount);
        Assert.Equal(visual ? 0 : 2, count.OmittedCount);
        Assert.Equal(visual ? 2 : 0, count.UnassessedCount);
        Assert.Equal(new ulong[] { 10, 11 }, report.SourceUnits
            .Where(unit => unit.Kind == IWorkSourceUnitKind.UnsupportedObject)
            .Select(unit => unit.Identity.RecordIdentifier).OrderBy(id => id));
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier is 3 or 99);
        Assert.All(report.SourceUnits.Where(unit => unit.Kind == IWorkSourceUnitKind.UnsupportedObject),
            unit => Assert.Equal(unit.Identity.RecordIdentifier == 10 ? 9999u : 9998u, unit.Identity.MessageType));
    }
}
