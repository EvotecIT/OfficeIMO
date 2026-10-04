using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 2, false)]
    [InlineData(IWorkDocumentKind.Pages, 2, true)]
    [InlineData(IWorkDocumentKind.Pages, 4, false)]
    [InlineData(IWorkDocumentKind.Pages, 4, true)]
    [InlineData(IWorkDocumentKind.Numbers, 2, false)]
    [InlineData(IWorkDocumentKind.Numbers, 2, true)]
    [InlineData(IWorkDocumentKind.Keynote, 2, false)]
    [InlineData(IWorkDocumentKind.Keynote, 2, true)]
    [InlineData(IWorkDocumentKind.Keynote, 4, false)]
    [InlineData(IWorkDocumentKind.Keynote, 4, true)]
    public void Selected_text_shapes_retain_missing_storage_evidence(IWorkDocumentKind kind, int field, bool visual) {
        using MemoryStream package = TextReferencePackage(kind, ReferenceField(field, 999));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual);
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), 10, field.ToString(), 999);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 999);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Declared_alternate_text_storage_is_assessed_even_when_the_other_field_resolves(IWorkDocumentKind kind) {
        using MemoryStream package = TextReferencePackage(kind,
            Message(ReferenceField(2, 11), ReferenceField(4, 999)),
            ArchiveRecord(11, 2001, Message(StringField(3, "Text"))));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, kind).SourceReferenceIssues), 10, "4", 999);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Malformed_text_storage_reference_sets_preserve_readable_siblings(IWorkDocumentKind kind) {
        using MemoryStream package = TextReferencePackage(kind,
            Message(ReferenceField(2, 11), BytesField(2, new byte[] { 0x80 })),
            ArchiveRecord(11, 2001, Message(StringField(3, "Text"))));
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Equal(2, report.SourceReferenceIssues.Count);
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, report.SourceReferenceIssues[0].Kind);
        Assert.Equal(11ul, report.SourceReferenceIssues[0].TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, report.SourceReferenceIssues[1].Kind);
        Assert.Null(report.SourceReferenceIssues[1].TargetIdentifier);
    }

    [Theory]
    [InlineData(7u)]
    [InlineData(2011u)]
    public void Selected_keynote_nested_storage_retains_its_owner_and_path(uint nativeType) {
        using MemoryStream package = TextReferencePackage(IWorkDocumentKind.Keynote,
            BytesField(1, Message(ReferenceField(2, 999))), nativeType: nativeType);
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Keynote).SourceReferenceIssues),
            10, "1/2", 999);
    }

    [Fact]
    public void Resolved_keynote_direct_storage_does_not_assess_unused_nested_storage() {
        using MemoryStream package = TextReferencePackage(IWorkDocumentKind.Keynote,
            Message(ReferenceField(2, 11), BytesField(1, Message(ReferenceField(2, 999)))),
            ArchiveRecord(11, 2001, Message(StringField(3, "Text"))));
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Keynote).SourceReferenceIssues);
    }

    [Fact]
    public void Pages_section_table_failures_retain_the_body_nested_path() {
        using MemoryStream package = CreatePagesPackageWithMissingSection();
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues),
            2, "17/1[1]/2", 3);
    }

    [Theory]
    [InlineData(23)]
    [InlineData(24)]
    [InlineData(25)]
    public void Declared_pages_header_templates_retain_missing_archive_evidence(int field) {
        using MemoryStream package = HeaderReferencePackage(ReferenceField(field, 999), Message());
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues),
            3, field.ToString(), 999);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void Shared_pages_template_archives_count_each_physical_reference_field_once(int field) {
        byte[] templates = Message(ReferenceField(23, 4), ReferenceField(24, 4), ReferenceField(25, 4));
        using MemoryStream package = HeaderReferencePackage(templates,
            Message(ReferenceField(field, 999), ReferenceField(field, 999)), twoSections: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 2 });
        IWorkPagesProjection projection = source.ReadPages();
        Assert.Equal(2, projection.SourceReferenceIssues.Count);
        Assert.Equal(new[] { 1, 2 }, projection.SourceReferenceIssues.Select(issue => issue.ReferenceIndex));
        Assert.All(projection.SourceReferenceIssues, issue => AssertMissingReference(issue, 4, field.ToString(), 999));
        Assert.Equal(2, projection.Sections.Count);
    }

    [Fact]
    public void Unselected_pages_header_archives_remain_outside_reference_evidence() {
        using MemoryStream package = HeaderReferencePackage(Message(),
            Message(ReferenceField(1, 999), ReferenceField(2, 998)));
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues);
    }

    [Fact]
    public void Header_reference_limit_failure_is_not_swallowed_as_a_partial_projection() {
        using MemoryStream package = HeaderReferencePackage(ReferenceField(25, 4),
            Message(ReferenceField(1, 999), ReferenceField(2, 998)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        Assert.Throws<InvalidDataException>(() => source.ReadPages());
    }

    [Fact]
    public void Presenter_note_storage_reference_limit_failure_is_not_swallowed_as_a_partial_projection() {
        using MemoryStream package = CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(27, 5))),
            ArchiveRecord(5, 15, Message(ReferenceField(1, 999), ReferenceField(1, 998)))))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Keynote,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        Assert.Throws<InvalidDataException>(() => source.ReadKeynote());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Presenter_note_and_note_storage_failures_retain_separate_owner_paths(bool storage, bool visual) {
        using MemoryStream package = CreatePackage(("Index/Slide.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(27, storage ? 5ul : 999ul))),
            storage ? ArchiveRecord(5, 15, Message(ReferenceField(1, 999))) : Array.Empty<byte>()))),
            ("preview.png", ValidPreviewPng()));
        AssertMissingReference(Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Keynote, visual).SourceReferenceIssues),
            storage ? 5ul : 4ul, storage ? "1" : "27", 999);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Graph_and_text_reference_evidence_share_the_limit_and_propagate_failure(IWorkDocumentKind kind) {
        using MemoryStream package = TextReferencePackage(kind, ReferenceField(2, 999), graphFailure: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => {
            if (kind == IWorkDocumentKind.Pages) source.ReadPages();
            else if (kind == IWorkDocumentKind.Numbers) source.ReadNumbers();
            else source.ReadKeynote();
        });
        Assert.Contains("reference issues", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static void AssertMissingReference(IWorkSourceReferenceIssue issue, ulong owner, string path, ulong target) {
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(target, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.MissingTarget, issue.Kind);
    }

    private static MemoryStream TextReferencePackage(IWorkDocumentKind kind, byte[] shape,
        byte[]? additionalRecords = null, uint nativeType = 2011, bool graphFailure = false) {
        int drawableField = kind == IWorkDocumentKind.Pages ? 1 : kind == IWorkDocumentKind.Numbers ? 2 : 7;
        byte[] drawables = Message(ReferenceField(drawableField, 10),
            graphFailure ? ReferenceField(drawableField, 998) : Array.Empty<byte>());
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(ArchiveRecord(1, 10000,
                    Message(ReferenceField(4, 2), ReferenceField(20, 3))),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"))), ArchiveRecord(3, 10020, drawables)),
            IWorkDocumentKind.Numbers => Message(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), drawables))),
            _ => Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))), ArchiveRecord(4, 5, drawables))
        };
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, nativeType, shape), additionalRecords ?? Array.Empty<byte>()))),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream HeaderReferencePackage(byte[] templates, byte[] archive, bool twoSections = false) {
        byte[] sectionEntry = BytesField(1, Message(ReferenceField(2, 3)));
        byte[] sectionTable = Message(sectionEntry, twoSections ? sectionEntry : Array.Empty<byte>());
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
            ArchiveRecord(2, 2001, Message(StringField(3, twoSections ? "First\u0004Second" : "Body"),
                BytesField(17, sectionTable))),
            ArchiveRecord(3, 10011, templates), ArchiveRecord(4, 10143, archive)))),
            ("preview.png", ValidPreviewPng()));
    }
}
