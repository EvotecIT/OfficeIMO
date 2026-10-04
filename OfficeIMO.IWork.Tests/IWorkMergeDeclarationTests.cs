using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> MalformedMergePaths() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (int depth in new[] { 0, 1, 2 }) yield return new object[] { kind, depth };
    }

    [Theory]
    [MemberData(nameof(MalformedMergePaths))]
    public void Unreadable_merge_declaration_retains_owner_and_physical_path(IWorkDocumentKind kind, int depth) {
        byte[] merge = depth switch {
            0 => BytesField(47, new byte[] { 0x80 }),
            1 => BytesField(47, BytesField(2, new byte[] { 0x80 })),
            _ => MergeTestOwner(BytesField(3, new byte[] { 0x80 }), MergeTestPair(3, 0, 3, 1))
        };
        using MemoryStream package = MergeTestPackage(kind, merge);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.Empty(table.MergedRanges);
        package.Position = 0;
        IWorkSourceDeclarationIssue issue = Assert.Single(ConvertUnitReport(package, kind).SourceDeclarationIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal(depth == 0 ? "47" : depth == 1 ? "47/2" : "47/2/3[1]", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Conflicting_merge_declarations_retire_both_ranges_and_keep_disjoint_range(IWorkDocumentKind kind) {
        using MemoryStream package = MergeTestPackage(kind, ConflictingMergeTestOwner());
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        IWorkTableMergeRange merge = Assert.Single(table.MergedRanges);
        Assert.Equal((4, 1, 4, 2), (merge.FirstRow, merge.FirstColumn, merge.LastRow, merge.LastColumn));
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, issue => issue.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && issue.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Equal(new[] { "47/2/3[1]/2", "47/2/3[2]/2" }, report.SourceDeclarationIssues.Select(issue => issue.FieldPath));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Disjoint_merge_recovery_allows_partial_destination_save_and_reopen(IWorkDocumentKind kind, bool unreadableRange) {
        byte[] owner = unreadableRange ? MergeTestOwner(BytesField(3, new byte[] { 0x80 }), MergeTestPair(3, 0, 3, 1))
            : ConflictingMergeTestOwner();
        using MemoryStream package = MergeTestPackage(kind, owner);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options);
            Assert.False(partial.IsVisualFallback);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            Assert.Equal("42", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Equal(unreadableRange ? 4 : 3, reopened.Tables[0].Rows[3].Cells.Count);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options);
            Assert.False(partial.IsVisualFallback);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(saved);
            Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            if (unreadableRange) Assert.Empty(reopened.Sheets[0].GetMergedRanges());
            else Assert.Equal("A4:B4", Assert.Single(reopened.Sheets[0].GetMergedRanges()).A1Range);
        } else {
            using var automatic = source.ToPowerPointPresentationResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options);
            Assert.False(partial.IsVisualFallback);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            Assert.Equal("42", reopened.Slides[0].Tables.First().GetCell(0, 0).Text);
            Assert.Equal(unreadableRange ? 1 : 2, reopened.Slides[0].Tables.First().GetCell(3, 0).Merge.columns);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Known_out_of_bounds_rectangle_retires_only_intersecting_merges_without_exporting_a_clip(bool intersects) {
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, MergeTestOwner(
            MergeTestPair(0, 0, 1, 1), MergeTestPair(3, 0, 7, intersects ? 3ul : 1ul), MergeTestPair(3, 2, 3, 3)));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(intersects ? 1 : 2, table.MergedRanges.Count);
        Assert.Contains(table.MergedRanges, range => range.FirstRow == 1 && range.LastRow == 2);
        Assert.DoesNotContain(table.MergedRanges, range => range.FirstRow == 4 && range.FirstColumn == 1);
        Assert.Equal(intersects ? 2 : 1, projection.SourceDeclarationIssues.Count);
        Assert.Equal("47/2/3[2]/2", projection.SourceDeclarationIssues[0].FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, projection.SourceDeclarationIssues[0].Kind);
    }

    [Fact]
    public void Identical_valid_ranges_normalize_without_disabling_editable_content() {
        byte[] pair = MergeTestPair(3, 0, 3, 1);
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, MergeTestOwner(pair, pair, pair));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).MergedRanges);
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceDeclarationIssues);
    }

    [Fact]
    public void All_physical_aliases_of_conflicting_ranges_retain_evidence() {
        byte[] pair = MergeTestPair(0, 0, 1, 1);
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers,
            MergeTestOwner(pair, MergeTestPair(1, 1, 2, 2), pair));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.Empty(Assert.Single(Assert.Single(projection.Sheets).Tables).MergedRanges);
        Assert.Equal(new[] { "47/2/3[1]/2", "47/2/3[2]/2", "47/2/3[3]/2" },
            projection.SourceDeclarationIssues.Select(issue => issue.FieldPath).OrderBy(path => path));
    }

    [Theory]
    [InlineData("missing", 0, IWorkSourceDeclarationIssueKind.RejectedMessageSet)]
    [InlineData("wire", 1, IWorkSourceDeclarationIssueKind.RejectedMessageSet)]
    [InlineData("malformed", 1, IWorkSourceDeclarationIssueKind.MalformedMessage)]
    [InlineData("unsupported", 1, IWorkSourceDeclarationIssueKind.RejectedMessageSet)]
    public void Unresolved_merge_expression_retains_known_field_count_and_never_trusts_siblings(string defect,
        int count, IWorkSourceDeclarationIssueKind kind) {
        byte[] formulaField = defect switch {
            "missing" => Message(), "wire" => VarintField(2, 0),
            "malformed" => BytesField(2, new byte[] { 0x80 }), _ => BytesField(2, FormulaConstant(1d))
        };
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers,
            MergeTestOwner(BytesField(3, formulaField), MergeTestPair(3, 0, 3, 1)));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.Empty(Assert.Single(Assert.Single(projection.Sheets).Tables).MergedRanges);
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal("47/2/3[1]/2", issue.FieldPath);
        Assert.Equal(count, issue.DeclaredValueCount);
        Assert.Equal(kind, issue.Kind);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void Configured_merge_message_field_limits_remain_fatal_at_every_envelope(int depth) {
        byte[] fields = Message(Enumerable.Range(0, 9).Select(_ => VarintField(1, 0)).ToArray());
        byte[] owner = depth switch {
            0 => BytesField(47, fields), 1 => BytesField(47, BytesField(2, fields)),
            2 => MergeTestOwner(BytesField(3, fields)), _ => MergeTestOwner(BytesField(3, BytesField(2, fields)))
        };
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, owner);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Merge_formula_depth_is_preserved_through_physical_entry_parsing() {
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, MergeTestOwner(MergeTestPair(3, 0, 3, 1)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProtobufDepth = 6 });
        Assert.Contains("depth", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Unreadable_pair_does_not_hide_a_later_formula_node_budget_failure() {
        byte[] formula = BytesField(1, Message(BytesField(1, Message(VarintField(1, 17), DoubleField(4, 1))),
            BytesField(1, new byte[] { 0x80 })));
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers,
            MergeTestOwner(BytesField(3, new byte[] { 0x80 }), BytesField(3, BytesField(2, formula))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumFormulaNodes = 1 });
        Assert.Contains("syntax-node limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Merge_evidence_budget_is_fatal_cumulative_and_snapshotted() {
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, ConflictingMergeTestOwner());
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Shared_model_conflicts_report_each_physical_path_once() {
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, ConflictingMergeTestOwner(), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 2 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Equal(2, projection.SourceDeclarationIssues.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Repeated_merge_owner_or_store_retains_actual_outer_count_without_assessing_nested_pairs(bool store) {
        byte[] pairs = MergeTestPair(3, 0, 3, 1);
        byte[] owner = store ? BytesField(47, Message(BytesField(2, pairs), BytesField(2, pairs)))
            : Message(MergeTestOwner(pairs), MergeTestOwner(pairs));
        using MemoryStream package = MergeTestPackage(IWorkDocumentKind.Numbers, owner);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.Empty(Assert.Single(Assert.Single(projection.Sheets).Tables).MergedRanges);
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(store ? "47/2" : "47", issue.FieldPath);
        Assert.Equal(2, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.RejectedMessageSet, issue.Kind);
    }

    private static byte[] ConflictingMergeTestOwner() => MergeTestOwner(MergeTestPair(0, 0, 1, 1),
        MergeTestPair(1, 1, 2, 2), MergeTestPair(3, 0, 3, 1));

    private static byte[] MergeTestOwner(params byte[][] pairs) => BytesField(47, BytesField(2, Message(pairs)));

    private static byte[] MergeTestPair(ulong firstRow, ulong firstColumn, ulong lastRow, ulong lastColumn) {
        byte[] tract = Message(BytesField(4, Message(VarintField(1, firstRow), VarintField(2, lastRow))),
            BytesField(3, Message(VarintField(1, firstColumn), VarintField(2, lastColumn))));
        byte[] node = Message(VarintField(1, 67), BytesField(40, tract));
        return BytesField(3, BytesField(2, BytesField(1, BytesField(1, node))));
    }

    private static MemoryStream MergeTestPackage(IWorkDocumentKind kind, byte[] merge, bool repeatModel = false) {
        byte[] store = BytesField(3, BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))));
        byte[] model = Message(BytesField(4, store), VarintField(6, 4), VarintField(7, 4), merge);
        return TableDependencyPackage(kind, Message(), modelPayload: model, rows: 4, columns: 4, repeatModel: repeatModel);
    }
}
