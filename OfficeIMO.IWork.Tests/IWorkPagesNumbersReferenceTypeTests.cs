using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 2)]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 2)]
    public void Wrong_type_pages_numbers_shape_storage_retains_selected_reference_identity(IWorkDocumentKind kind, int field) {
        using var package = TextReferencePackage(kind, ReferenceField(field, 11), ArchiveRecord(11, 15, Message()));
        AssertWrongTypeReference(ConvertUnitReport(package, kind), 10, field.ToString(), 11);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 1)]
    public void Wrong_type_pages_body_or_numbers_sheet_is_not_reported_as_missing(IWorkDocumentKind kind, int field) {
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, kind == IWorkDocumentKind.Pages ? 10000u : 1u, ReferenceField(field, 2)),
            ArchiveRecord(2, 2011, Message())))), ("preview.png", ValidPreviewPng()));
        AssertWrongTypeReference(ConvertUnitReport(package, kind), 1, field.ToString(), 2);
    }

    [Theory]
    [InlineData(0, 2ul, "17/1[1]/2", 3ul)]
    [InlineData(1, 3ul, "25", 4ul)]
    [InlineData(2, 4ul, "1", 5ul)]
    [InlineData(3, 4ul, "2", 5ul)]
    public void Wrong_type_pages_section_dependencies_retain_the_rejected_owner_path(int variant, ulong owner, string path, ulong target) {
        using var package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, ReferenceField(4, 2)),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"), BytesField(17,
                BytesField(1, ReferenceField(2, 3))))),
            ArchiveRecord(3, variant == 0 ? 2011u : 10011u, ReferenceField(25, 4)),
            ArchiveRecord(4, variant == 1 ? 2011u : 10143u, ReferenceField(variant == 3 ? 2 : 1, 5)),
            ArchiveRecord(5, 2011, Message())))), ("preview.png", ValidPreviewPng()));
        AssertWrongTypeReference(ConvertUnitReport(package, IWorkDocumentKind.Pages), owner, path, target);
    }

    private static void AssertWrongTypeReference(IWorkConversionReport report, ulong owner, string path, ulong target) {
        var issue = Assert.Single(report.SourceReferenceIssues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(target, issue.TargetIdentifier);
    }
}
