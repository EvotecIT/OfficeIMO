using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Missing_selected_drawable_references_retain_occurrences_without_inventing_objects(
        IWorkDocumentKind kind, bool visual) {
        using MemoryStream package = ReferenceIssuePackage(kind,
            field => Message(ReferenceField(field, 999), ReferenceField(field, 999)));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual);
        Assert.Equal(2, report.SourceReferenceIssues.Count);
        Assert.Equal(new[] { 1, 2 }, report.SourceReferenceIssues.Select(issue => issue.ReferenceIndex));
        Assert.All(report.SourceReferenceIssues, issue => {
            Assert.Equal(IWorkSourceReferenceIssueKind.MissingTarget, issue.Kind);
            Assert.Equal(999ul, issue.TargetIdentifier);
            Assert.Equal(kind == IWorkDocumentKind.Pages ? 3ul : kind == IWorkDocumentKind.Numbers ? 2ul : 4ul,
                issue.Owner.RecordIdentifier);
            Assert.Equal(kind == IWorkDocumentKind.Pages ? "1" : kind == IWorkDocumentKind.Numbers ? "2" : "6", issue.FieldPath);
        });
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier is 999 or 99);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Rejected_drawable_reference_sets_distinguish_malformed_values_from_readable_siblings(IWorkDocumentKind kind) {
        using MemoryStream package = ReferenceIssuePackage(kind,
            field => Message(ReferenceField(field, 10), BytesField(field, new byte[] { 0x80 }), VarintField(field, 10)));
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Equal(3, report.SourceReferenceIssues.Count);
        IWorkSourceReferenceIssue sibling = report.SourceReferenceIssues[0];
        Assert.Equal(10ul, sibling.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, sibling.Kind);
        Assert.All(report.SourceReferenceIssues.Skip(1), issue => {
            Assert.Null(issue.TargetIdentifier);
            Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, issue.Kind);
        });
    }

    [Fact]
    public void Missing_sheet_and_slide_tree_references_keep_their_selected_owner_paths() {
        using MemoryStream numbers = CreatePackage(("Index/Document.iwa", FrameIwa(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 999))))), ("preview.png", ValidPreviewPng()));
        IWorkSourceReferenceIssue sheet = Assert.Single(ConvertUnitReport(numbers, IWorkDocumentKind.Numbers).SourceReferenceIssues);
        Assert.Equal(1ul, sheet.Owner.RecordIdentifier); Assert.Equal("1", sheet.FieldPath);
        using MemoryStream keynote = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 999))))))),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceReferenceIssue node = Assert.Single(ConvertUnitReport(keynote, IWorkDocumentKind.Keynote).SourceReferenceIssues);
        Assert.Equal(2ul, node.Owner.RecordIdentifier); Assert.Equal("3/2", node.FieldPath);
        Assert.Equal(999ul, node.TargetIdentifier);
    }

    [Fact]
    public void Rejected_single_reference_fields_preserve_readable_targets_without_claiming_they_are_missing() {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), ReferenceField(4, 2))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body")))))), ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Pages);
        Assert.Equal(2, report.SourceReferenceIssues.Count);
        Assert.All(report.SourceReferenceIssues, issue => {
            Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, issue.Kind);
            Assert.Equal(2ul, issue.TargetIdentifier); Assert.Equal("4", issue.FieldPath);
        });
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 2);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 2)]
    [InlineData(IWorkDocumentKind.Keynote, 3)]
    public void Required_graph_reference_failures_survive_early_projection_returns(IWorkDocumentKind kind, int path) {
        byte[] records = path == 3
            ? Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 999))))
            : ArchiveRecord(1, kind == IWorkDocumentKind.Pages ? 10000u : 1u,
                Message(ReferenceField(path, 999)));
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceReferenceIssue issue = Assert.Single(ConvertUnitReport(package, kind).SourceReferenceIssues);
        Assert.Equal(path == 3 ? 3ul : 1ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path == 4 ? "4" : "2", issue.FieldPath);
        Assert.Equal(IWorkSourceReferenceIssueKind.MissingTarget, issue.Kind);
        Assert.Equal(999ul, issue.TargetIdentifier);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Source_reference_evidence_enforces_the_captured_budget(IWorkDocumentKind kind) {
        using MemoryStream package = ReferenceIssuePackage(kind,
            field => Message(ReferenceField(field, 999), ReferenceField(field, 999)));
        var options = new IWorkReadOptions { MaximumSourceReferenceIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind, options);
        options.MaximumSourceReferenceIssues = 100;
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => {
            if (kind == IWorkDocumentKind.Pages) source.ReadPages();
            else if (kind == IWorkDocumentKind.Numbers) source.ReadNumbers();
            else source.ReadKeynote();
        });
        Assert.Contains("reference issues", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unambiguous_reference_sets_report_only_failed_occurrences(IWorkDocumentKind kind) {
        using MemoryStream package = ReferenceIssuePackage(kind, field => Message(
            ReferenceField(field, 10), BytesField(field, Message(VarintField(1, 10), VarintField(1, 11))),
            ReferenceField(field, 999)));
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Equal(new[] { 2, 3 }, report.SourceReferenceIssues.Select(issue => issue.ReferenceIndex));
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, report.SourceReferenceIssues[0].Kind);
        Assert.Null(report.SourceReferenceIssues[0].TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.MissingTarget, report.SourceReferenceIssues[1].Kind);
        Assert.Equal(999ul, report.SourceReferenceIssues[1].TargetIdentifier);
        Assert.Contains(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 10);
    }

    [Fact]
    public void Reference_evidence_budget_is_shared_across_selected_sheets() {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2), ReferenceField(1, 3))),
            ArchiveRecord(2, 2, Message(ReferenceField(2, 999), ReferenceField(2, 998))),
            ArchiveRecord(3, 2, Message(ReferenceField(2, 997), ReferenceField(2, 996)))))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 3 });
        Assert.Throws<InvalidDataException>(() => source.ReadNumbers());
    }

    [Fact]
    public void Floating_canvas_reference_issues_keep_the_nested_field_path() {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), ReferenceField(3, 3))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10021, Message(BytesField(1, Message(BytesField(2,
                Message(ReferenceField(1, 999)))))))))), ("preview.png", ValidPreviewPng()));
        IWorkSourceReferenceIssue issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Pages).SourceReferenceIssues);
        Assert.Equal(3ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1[1]/2[1]/1", issue.FieldPath);
        Assert.Equal(999ul, issue.TargetIdentifier);
    }

    private static MemoryStream ReferenceIssuePackage(IWorkDocumentKind kind, Func<int, byte[]> references) {
        byte[] mapped = references(kind == IWorkDocumentKind.Pages ? 1 : kind == IWorkDocumentKind.Numbers ? 2 : 6);
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), ReferenceField(20, 3)), new ulong[] { 2, 3 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"))), ArchiveRecord(3, 10020, mapped)),
            IWorkDocumentKind.Numbers => Message(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), mapped))),
            _ => Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))), ArchiveRecord(4, 5, mapped))
        };
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, 9999, Message()), ArchiveRecord(99, 2, Message(ReferenceField(2, 998)))))),
            ("preview.png", ValidPreviewPng()));
    }
}
