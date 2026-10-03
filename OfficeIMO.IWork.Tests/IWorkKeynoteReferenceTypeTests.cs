using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Wrong_type_presenter_note_targets_retain_owner_path_and_target_identity(bool storage, bool visual) {
        using var package = CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(27, 5))),
            ArchiveRecord(5, storage ? 15u : 2011u, Message(ReferenceField(1, 6))),
            ArchiveRecord(6, storage ? 2011u : 2001u, Message(StringField(3, "Unselected text")))))),
            ("preview.png", ValidPreviewPng()));
        var report = ConvertUnitReport(package, IWorkDocumentKind.Keynote, visual);
        var issue = Assert.Single(report.SourceReferenceIssues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(storage ? 5ul : 4ul, issue.Owner.RecordIdentifier);
        Assert.Equal(storage ? "1" : "27", issue.FieldPath);
        Assert.Equal(storage ? 6ul : 5ul, issue.TargetIdentifier);
    }
    [Theory]
    [InlineData(0, 1ul, "2", 2ul)]
    [InlineData(1, 2ul, "3/2", 3ul)]
    [InlineData(2, 3ul, "2", 4ul)]
    public void Wrong_type_keynote_graph_targets_retain_selected_reference_evidence(int variant, ulong owner, string path, ulong target) {
        using var package = CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, variant == 0 ? 2011u : 2u, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, variant == 1 ? 2011u : 4u, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, variant == 2 ? 2011u : 5u, Message())))),
            ("preview.png", ValidPreviewPng()));
        var issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Keynote).SourceReferenceIssues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(target, issue.TargetIdentifier);
    }

    [Theory]
    [InlineData(2, "2")]
    [InlineData(4, "4")]
    [InlineData(1, "1/2")]
    public void Wrong_type_keynote_drawable_storage_retains_its_direct_or_nested_path(int field, string path) {
        byte[] shape = field == 1 ? BytesField(1, ReferenceField(2, 11)) : ReferenceField(field, 11);
        using var package = TextReferencePackage(IWorkDocumentKind.Keynote, shape,
            ArchiveRecord(11, 15, Message()));
        var issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Keynote).SourceReferenceIssues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(10ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(11ul, issue.TargetIdentifier);
    }

}
