using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> UnreadableAttributeCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            for (int defect = 0; defect < 6; defect++)
                foreach (bool visual in new[] { false, true }) yield return new object[] { kind, defect, visual };
    }

    [Theory]
    [MemberData(nameof(UnreadableAttributeCases))]
    public void Undecoded_attributes_retain_declaration_evidence_without_inventing_references(
        IWorkDocumentKind kind, int defect, bool visual) {
        byte[] attributes = defect switch {
            0 => BytesField(8, new byte[] { 0x80 }),
            1 => AttributeTable(8, new byte[] { 0x80 }),
            2 => BytesField(8, VarintField(1, 7)),
            3 => AttributeTable(8, AttributeEntry(50, ReferenceField(2, 999))),
            4 => Message(AttributeTable(8, AttributeEntry(0, ReferenceField(2, 999))), AttributeTable(8)),
            _ => BytesField(8, Message(VarintField(2, 7)))
        };
        using MemoryStream package = SelectedRichTextPackage(kind, attributes);
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(15ul, issue.Owner.RecordIdentifier);
        Assert.Equal(defect is 1 or 2 or 3 ? "8/1[1]" : "8", issue.FieldPath);
        Assert.Equal(defect == 4 ? 2 : 1, issue.DeclaredValueCount);
        Assert.Equal(defect == 3 ? IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata
            : defect >= 4 ? IWorkSourceDeclarationIssueKind.RejectedMessageSet
            : IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
        Assert.Empty(report.SourceReferenceIssues);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 999);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED").LossKind);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkSourceDeclarationIssue>)report.SourceDeclarationIssues).Clear());
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Reused_storage_declarations_are_reported_once_under_the_source_limit(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind, BytesField(8, new byte[] { 0x80 }), aliases: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        IReadOnlyList<IWorkSourceDeclarationIssue> issues = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().SourceDeclarationIssues,
            IWorkDocumentKind.Numbers => source.ReadNumbers().SourceDeclarationIssues,
            _ => source.ReadKeynote().SourceDeclarationIssues
        };
        Assert.Equal("8", Assert.Single(issues).FieldPath);
        Assert.Equal(15ul, issues[0].Owner.RecordIdentifier);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Declaration_limit_remains_fatal_and_uses_the_captured_options(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            Message(BytesField(5, new byte[] { 0x80 }), BytesField(8, new byte[] { 0x80 })));
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind)).Message);
    }

    public static IEnumerable<object[]> UnreadableTableCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            for (int defect = 0; defect < 5; defect++) yield return new object[] { kind, defect };
    }

    [Theory]
    [MemberData(nameof(UnreadableTableCases))]
    public void Undecoded_table_containers_retain_owner_paths_and_known_outer_counts(IWorkDocumentKind kind, int defect) {
        byte[] fields = defect switch {
            0 => BytesField(1, new byte[] { 0x80 }),
            1 => BytesField(3, new byte[] { 0x80 }),
            2 => BytesField(3, Message(VarintField(1, 1))),
            _ => Message()
        };
        byte[]? model = defect == 3 ? Message(BytesField(4, new byte[] { 0x80 }), VarintField(6, 1), VarintField(7, 1))
            : defect == 4 ? new byte[] { 0x80 } : null;
        using MemoryStream package = TableDependencyPackage(kind, fields, includeTile: defect == 0,
            modelPayload: model);
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal(new[] { "4/1", "4/3", "4/3/1", "4", "$" }[defect], issue.FieldPath);
        Assert.Equal(defect == 4 ? null : (int?)1, issue.DeclaredValueCount);
        Assert.Empty(report.SourceReferenceIssues);
    }

    [Fact]
    public void Malformed_sections_and_floating_canvas_entries_keep_their_selected_paths() {
        using MemoryStream sections = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"), BytesField(17, new byte[] { 0x80 })))))),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceDeclarationIssue section = Assert.Single(ConvertUnitReport(sections, IWorkDocumentKind.Pages).SourceDeclarationIssues);
        Assert.Equal(2ul, section.Owner.RecordIdentifier);
        Assert.Equal("17", section.FieldPath);
        using MemoryStream floating = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2), ReferenceField(3, 3))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10016, Message(BytesField(1, Message(VarintField(2, 99)))))))),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceDeclarationIssue canvas = Assert.Single(ConvertUnitReport(floating, IWorkDocumentKind.Pages).SourceDeclarationIssues);
        Assert.Equal(3ul, canvas.Owner.RecordIdentifier);
        Assert.Equal("1[1]/2", canvas.FieldPath);
        Assert.Equal(1, canvas.DeclaredValueCount);
    }

    [Fact]
    public void Malformed_keynote_tree_retains_evidence_through_the_early_return() {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, Message(BytesField(3, new byte[] { 0x80 })))))), ("preview.png", ValidPreviewPng()));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Keynote);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(2ul, issue.Owner.RecordIdentifier);
        Assert.Equal("3", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Empty(report.SourceReferenceIssues);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Malformed_selected_root_reports_unknown_nested_counts(IWorkDocumentKind kind) {
        using MemoryStream package = CreatePackage(("Index/Document.iwa", FrameIwa(
            ArchiveRecord(1, kind == IWorkDocumentKind.Pages ? 10000u : 1u, new byte[] { 0x80 }))),
            ("preview.png", ValidPreviewPng()));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        IWorkConversionReport report = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().CreateConversionReport(IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview),
            IWorkDocumentKind.Numbers => source.ReadNumbers().CreateConversionReport(IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview),
            _ => source.ReadKeynote().CreateConversionReport(IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview)
        };
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(1ul, issue.Owner.RecordIdentifier);
        Assert.Equal("$", issue.FieldPath);
        Assert.Null(issue.DeclaredValueCount);
    }

    [Fact]
    public void Declaration_options_are_positive_and_clone_preserves_the_limit() {
        Assert.Throws<ArgumentOutOfRangeException>(() => new IWorkReadOptions { MaximumSourceDeclarationIssues = 0 }.Clone());
        Assert.Equal(7, new IWorkReadOptions { MaximumSourceDeclarationIssues = 7 }.Clone().MaximumSourceDeclarationIssues);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Attribute_field_limit_is_not_downgraded_to_declaration_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(8, Enumerable.Range(0, 9).Select(_ => AttributeEntry(0, Message())).ToArray()));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind)).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_unreadable_style_parent_envelope_retains_declaration_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = SelectedRichTextPackage(kind,
            AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19))),
            additionalRecords: ArchiveRecord(19, 2021, Message(BytesField(1, new byte[] { 0x80 }))));
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(19ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Empty(report.SourceReferenceIssues);
    }

    [Fact]
    public void Declaration_evidence_does_not_require_retained_source_payloads() {
        using MemoryStream package = SelectedRichTextPackage(IWorkDocumentKind.Numbers,
            BytesField(8, new byte[] { 0x80 }));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { PreserveSourceRecords = false });
        IWorkConversionReport report = source.ReadNumbers().CreateConversionReport(
            IWorkProjectionKind.VisualFallback, source.PreferredRasterPreview);
        Assert.Empty(report.PreservedRecords);
        Assert.Equal(15ul, Assert.Single(report.SourceDeclarationIssues).Owner.RecordIdentifier);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }
}
