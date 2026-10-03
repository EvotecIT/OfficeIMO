using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> WrongTypeRichDependencies() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (int record in new[] { 11, 13, 14, 15 }) yield return new object[] { kind, record };
    }

    [Theory]
    [MemberData(nameof(WrongTypeRichDependencies))]
    public void Selected_rich_text_structural_links_retain_existing_wrong_type_target_evidence(IWorkDocumentKind kind, int record) {
        using var package = SelectedRichTextPackage(kind, wrongTypeRecord: record);
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        var issue = Assert.Single(report.SourceReferenceIssues);
        (ulong owner, string path) = record switch { 11 => (10ul, "2"), 13 => (11ul, "4/17"),
            14 => (13ul, "3[1]/9"), _ => (14ul, "1") };
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal((ulong)record, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(1, issue.ReferenceIndex);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            d => d.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == (ulong)record);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unselected_wrong_type_rich_wrappers_stay_outside_the_reference_inventory(IWorkDocumentKind kind) {
        using var package = SelectedRichTextPackage(kind, wrongTypeRecord: 17);
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        Assert.Equal("Value", Assert.Single(table.Cells).RichText!.PlainText);
        Assert.Empty(issues);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Wrong_type_sizing_buckets_and_tiles_retain_selected_reference_paths_and_numeric_values(int field) {
        byte[] declaration = field switch {
            1 => BytesField(1, ReferenceField(2, 13)),
            3 => BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 13)))),
            _ => ReferenceField(field, 13)
        };
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, declaration,
            includeTile: field != 3, additionalRecords: ArchiveRecord(13, 2021, Message()));
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        var issue = Assert.Single(issues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal(field == 1 ? "4/1/2" : field == 3 ? "4/3/1[1]/2" : "4/2", issue.FieldPath);
        Assert.Equal(13ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        if (field != 3) Assert.Equal(42d, Assert.Single(table.Cells).Value);
    }

    [Theory]
    [InlineData("wrongAttachment", 2ul, "9/1[1]/2", 20ul)]
    [InlineData("wrongDrawable", 20ul, "1", 30ul)]
    public void Wrong_type_inline_objects_keep_reference_evidence_without_discarding_healthy_siblings(
        string defect, ulong owner, string path, ulong target) {
        using var package = InlineImagePackage(defect);
        var projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();
        Assert.True(projection.Body.HasUnresolvedInlineObjects);
        Assert.False(projection.HasEditableContent);
        var issue = Assert.Single(projection.SourceReferenceIssues);
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(target, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal(31ul, Assert.Single(projection.Body.Paragraphs.SelectMany(p => p.Runs), run => run.InlineObject != null).InlineObject!.Drawable.RecordIdentifier);
    }

    [Theory]
    [InlineData(13, 11ul, "4/19")]
    [InlineData(14, 13ul, "3[1]/10")]
    [InlineData(15, 14ul, "3")]
    public void Selected_comment_dependencies_keep_values_and_classify_wrong_target_types(int record, ulong owner, string path) {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(19, 13),
            cellPayload: CommentCell(false), additionalRecords: Message(
                ArchiveRecord(13, record == 13 ? 2021u : 6005u, Message(VarintField(1, 10),
                    BytesField(3, Message(VarintField(1, 1), ReferenceField(10, 14))))),
                ArchiveRecord(14, record == 14 ? 2021u : 3056u, Message(StringField(1, "Review"),
                    BytesField(2, DoubleField(1, 42d)), ReferenceField(3, 15))),
                ArchiveRecord(15, record == 15 ? 2021u : 212u, StringField(1, "Reviewer"))));
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        var cell = Assert.Single(table.Cells);
        Assert.Equal(42d, cell.Value);
        Assert.Null(cell.Comment);
        Assert.Equal(IWorkCellUnsupportedFeatures.Comment, cell.UnsupportedFeatures);
        var issue = Assert.Single(issues);
        Assert.Equal(owner, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal((ulong)record, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
    }
    [Fact]
    public void Wrong_type_number_format_catalog_preserves_raw_values_without_claiming_format_reconstruction() {
        using var package = NumberFormatPackage(IWorkDocumentKind.Numbers, NumericFormat(258, 2), catalogType: 2021);
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        var cell = Assert.Single(table.Cells);
        Assert.Equal(0.5d, cell.Value);
        Assert.Null(cell.NumberFormat);
        var issue = Assert.Single(issues);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Equal("4/22", issue.FieldPath);
        Assert.Equal(13ul, issue.TargetIdentifier);
    }

    [Fact]
    public void Unselected_lazy_cell_catalogs_do_not_assess_wrong_type_targets() {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(5, 13), ReferenceField(19, 13), ReferenceField(22, 13)),
            additionalRecords: ArchiveRecord(13, 2021, Message()));
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        Assert.Empty(issues);
    }

    [Fact]
    public void Wrong_type_row_bucket_invalidates_its_axis_preserves_column_sizes_and_charges_healthy_siblings() {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(BytesField(1, Message(ReferenceField(2, 13), ReferenceField(2, 14))), ReferenceField(2, 15)),
            additionalRecords: Message(ArchiveRecord(13, 6006, BytesField(2, Message(VarintField(1, 0), FloatField(2, 22), VarintField(3, 0)))),
                ArchiveRecord(14, 2021, Message()),
                ArchiveRecord(15, 6006, BytesField(2, Message(VarintField(1, 0), FloatField(2, 44), VarintField(3, 0))))));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceReferenceIssues = 2 });
        var projection = source.ReadNumbers();
        var table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Empty(table.RowHeights);
        Assert.Equal(44d, table.ColumnWidths[1]);
        var issue = Assert.Single(projection.SourceReferenceIssues);
        Assert.Equal(2, issue.ReferenceIndex);
        Assert.Equal(14ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Contains(projection.SourceDeclarationIssues, declaration => declaration.Owner.RecordIdentifier == 14
            && declaration.FieldPath == "$" && declaration.Kind == IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        package.Position = 0;
        var bounded = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceReferenceIssues = 1 });
        Assert.Contains("source reference issues", Assert.Throws<InvalidDataException>(() => bounded.ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(14)]
    [InlineData(15)]
    public void Shared_wrong_type_rich_dependencies_report_each_physical_link_once(int record) {
        using var package = SelectedRichTextPackage(IWorkDocumentKind.Numbers, aliases: true, wrongTypeRecord: record);
        int physicalLinks = record == 14 ? 2 : 1;
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceReferenceIssues = physicalLinks }), IWorkDocumentKind.Numbers);
        Assert.Equal(2, table.Cells.Count);
        Assert.Equal(physicalLinks, issues.Count);
        Assert.All(issues, issue => {
            Assert.Equal((ulong)record, issue.TargetIdentifier);
            Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        });
        Assert.Equal(record == 14 ? new[] { "3[1]/9", "3[2]/9" } : new[] { "1" }, issues.Select(issue => issue.FieldPath));
    }

    [Fact]
    public void Rejected_row_bucket_sets_keep_wrong_type_readable_siblings_classified_as_unselected() {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            BytesField(1, Message(ReferenceField(2, 13), VarintField(2, 13))),
            additionalRecords: ArchiveRecord(13, 2021, Message()));
        var (_, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.Equal(2, issues.Count);
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, issues[0].Kind);
        Assert.Equal(13ul, issues[0].TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, issues[1].Kind);
    }

    [Fact]
    public void Selected_wrong_type_cell_style_catalog_keeps_numeric_values_and_qualified_role_defaults() {
        using var package = RoleFillPackage(IWorkDocumentKind.Numbers, "wrong-catalog-target-type");
        var (table, issues) = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.Equal(42d, table.GetCell(3, 2)!.Value);
        Assert.Equal("0000FF", table.GetFill(3, 2)!.Color!.RgbHex);
        var issue = Assert.Single(issues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal("4/5", issue.FieldPath);
        Assert.Equal(13ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Wrong_type_style_parent_at_the_strict_depth_limit_remains_recoverable_reference_evidence(IWorkDocumentKind kind) {
        using var package = AttributeReferencePackage(kind, AttributeTable(5, AttributeEntry(0, ReferenceField(2, 10))),
            Message(ArchiveRecord(10, 2022, Message(BytesField(1, ReferenceField(3, 12)), BytesField(11, FloatField(3, 18)))),
                ArchiveRecord(12, 2021, Message())));
        IWorkConversionReport report = ConvertUnitReport(package, kind,
            readOptions: new IWorkReadOptions { MaximumTextStyleInheritanceDepth = 1 });
        var issue = Assert.Single(report.SourceReferenceIssues);
        Assert.Equal(10ul, issue.Owner.RecordIdentifier);
        Assert.Equal("1/3", issue.FieldPath);
        Assert.Equal(12ul, issue.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.UnexpectedTargetType, issue.Kind);
        Assert.Empty(report.SourceDeclarationIssues);
    }

}
