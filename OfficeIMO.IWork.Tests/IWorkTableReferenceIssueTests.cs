using OfficeIMO.IWork;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> TableDependencyCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages,
                IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote }) {
            foreach (int field in new[] { 1, 2, 3, 4, 6 }) yield return new object[] { kind, field };
        }
    }

    [Theory]
    [MemberData(nameof(TableDependencyCases))]
    public void Selected_table_dependencies_retain_missing_target_evidence(IWorkDocumentKind kind, int field) {
        byte[] declaration = field switch {
            1 => BytesField(1, Message(ReferenceField(2, 999))),
            3 => BytesField(3, Message(BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 999))))),
            _ => ReferenceField(field, 999)
        };
        using MemoryStream package = TableDependencyPackage(kind, declaration, includeTile: field != 3);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        (IWorkTable table, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(source, kind);
        string path = field switch { 1 => "4/1/2", 3 => "4/3/1[1]/2", _ => "4/" + field };
        AssertMissingReference(Assert.Single(issues), 11, path, 999);
        if (field != 3) Assert.Equal(42d, Assert.Single(table.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), 11, path, 999);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_REFERENCES_UNRESOLVED").LossKind);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 999);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, 4)]
    [InlineData(IWorkDocumentKind.Pages, 6)]
    [InlineData(IWorkDocumentKind.Numbers, 4)]
    [InlineData(IWorkDocumentKind.Numbers, 6)]
    [InlineData(IWorkDocumentKind.Keynote, 4)]
    [InlineData(IWorkDocumentKind.Keynote, 6)]
    public void Declared_missing_catalog_blocks_complete_editable_reconstruction(IWorkDocumentKind kind, int field) {
        using MemoryStream package = TableDependencyPackage(kind, ReferenceField(field, 999));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        bool complete = kind switch {
            IWorkDocumentKind.Pages => source.ReadPages().HasEditableContent,
            IWorkDocumentKind.Numbers => source.ReadNumbers().HasEditableContent,
            _ => source.ReadKeynote().HasEditableContent
        };
        Assert.False(complete);
    }

    [Theory]
    [InlineData(4)]
    [InlineData(6)]
    public void Rejected_catalog_reference_set_records_readable_sibling_and_malformed_occurrence(int field) {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(field, 13), VarintField(field, 13)),
            additionalRecords: ArchiveRecord(13, field == 4 ? 6200u : 6201u, Message()));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        Assert.Equal(2, projection.SourceReferenceIssues.Count);
        IWorkSourceReferenceIssue first = projection.SourceReferenceIssues[0];
        Assert.Equal("4/" + field, first.FieldPath);
        Assert.Equal(13ul, first.TargetIdentifier);
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, first.Kind);
        Assert.Equal(1, first.ReferenceIndex);
        IWorkSourceReferenceIssue second = projection.SourceReferenceIssues[1];
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, second.Kind);
        Assert.Equal(2, second.ReferenceIndex);
        Assert.Null(second.TargetIdentifier);
        Assert.Equal(42d, Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells).Value);
    }

    [Theory]
    [InlineData(4)]
    [InlineData(6)]
    public void Malformed_single_catalog_reference_is_not_an_absent_optional_catalog(int field) {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            BytesField(field, new byte[] { 0x80 }));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.False(projection.HasEditableContent);
        IWorkSourceReferenceIssue issue = Assert.Single(projection.SourceReferenceIssues);
        Assert.Equal("4/" + field, issue.FieldPath);
        Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, issue.Kind);
        Assert.Null(issue.TargetIdentifier);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unresolved_formula_catalog_preserves_cache_in_explicit_partial_conversion_and_reopen(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, ReferenceField(6, 999), formulaCell: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        (IWorkTable table, _) = ReadSelectedRichTable(source, kind);
        IWorkTableCell cell = Assert.Single(table.Cells);
        Assert.Equal(IWorkCellKind.Formula, cell.Kind);
        Assert.Equal(42d, cell.Value);
        Assert.False(cell.FormulaIsComplete);
        Assert.True(cell.CachedValueIsComplete);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkConversionReport report;
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options);
            Assert.False(partial.IsVisualFallback);
            report = partial.Report;
            partial.Value.Save(saved);
            saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            Assert.Equal("42", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options);
            Assert.False(partial.IsVisualFallback);
            report = partial.Report;
            partial.Value.Save(saved);
            saved.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(saved);
            Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue().Value);
        } else {
            using var automatic = source.ToPowerPointPresentationResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options);
            Assert.False(partial.IsVisualFallback);
            report = partial.Report;
            partial.Value.Save(saved);
            saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            Assert.Equal("42", reopened.Slides[0].Tables.First().GetCell(0, 0).Text);
        }
        Assert.True(report.IsPartialEditableReconstruction);
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), 11, "4/6", 999);
        IWorkFormulaCellStatus assessment = Assert.Single(report.FormulaCells);
        Assert.False(assessment.ExpressionIsComplete);
        Assert.Equal(IWorkFormulaCacheStatus.Complete, assessment.CacheStatus);
    }

    [Fact]
    public void Row_bucket_rejection_retains_readable_siblings_and_column_reference_evidence() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(BytesField(1, Message(ReferenceField(2, 13), VarintField(2, 13))),
                BytesField(2, new byte[] { 0x80 })),
            additionalRecords: ArchiveRecord(13, 6006, Message()));
        (_, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.Equal(new[] { "4/1/2", "4/1/2", "4/2" }, issues.Select(issue => issue.FieldPath));
        Assert.Equal(IWorkSourceReferenceIssueKind.RejectedReferenceSet, issues[0].Kind);
        Assert.Equal(13ul, issues[0].TargetIdentifier);
        Assert.All(issues.Skip(1), issue => Assert.Equal(IWorkSourceReferenceIssueKind.MalformedReference, issue.Kind));
    }

    [Fact]
    public void Tile_reference_path_keeps_physical_index_after_skipped_tile_entry() {
        byte[] tiles = BytesField(3, Message(BytesField(1, Message()),
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, 999)))));
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, tiles,
            includeTile: false, rows: 257);
        (_, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        AssertMissingReference(Assert.Single(issues), 11, "4/3/1[2]/2", 999);
    }

    [Fact]
    public void Invalid_tile_index_does_not_assess_its_unused_target() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            BytesField(3, Message(BytesField(1, Message(ReferenceField(2, 999))))), includeTile: false);
        (_, IReadOnlyList<IWorkSourceReferenceIssue> issues) = ReadSelectedRichTable(
            IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers);
        Assert.Empty(issues);
    }

    [Fact]
    public void Table_reference_budget_is_cumulative_across_sizing_and_catalog_fields() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(2, 999), ReferenceField(4, 998), ReferenceField(6, 997)));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 2 });
        Assert.Throws<InvalidDataException>(() => source.ReadNumbers());
    }

    [Fact]
    public void Repeated_selected_table_model_reports_each_physical_failed_field_once() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(2, 999), ReferenceField(4, 998)), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceReferenceIssues = 2 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Equal(2, projection.SourceReferenceIssues.Count);
        AssertMissingReference(projection.SourceReferenceIssues[0], 11, "4/2", 999);
        AssertMissingReference(projection.SourceReferenceIssues[1], 11, "4/4", 998);
    }

    [Fact]
    public void Optional_absent_catalogs_preserve_editable_numeric_cells() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message());
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(42d, Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells).Value);
    }

    private static MemoryStream TableDependencyPackage(IWorkDocumentKind kind, byte[] storeFields,
        bool includeTile = true, byte[]? additionalRecords = null, ulong rows = 1,
        bool formulaCell = false, bool repeatModel = false, byte[]? modelPayload = null) {
        byte[] cell = new byte[formulaCell ? 24 : 20]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, (1u << 1) | (formulaCell ? 1u << 9 : 0u));
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        byte[] roots = kind switch {
            IWorkDocumentKind.Pages => Message(
                ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2, 10 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body")))),
            IWorkDocumentKind.Numbers => Message(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
                ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10),
                    repeatModel ? ReferenceField(2, 20) : Message()))),
            _ => Message(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
                ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
                ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
                ArchiveRecord(4, 5, Message(ReferenceField(7, 10))))
        };
        byte[] store = Message(storeFields, includeTile ? BytesField(3, Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, 12))))) : Message());
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(roots,
            ArchiveRecord(10, 6000, Message(BytesField(1, GeometryDrawable(72, 72, 120, 40)), ReferenceField(2, 11))),
            repeatModel ? ArchiveRecord(20, 6000, Message(ReferenceField(2, 11))) : Message(),
            ArchiveRecord(11, 6001, modelPayload ?? Message(BytesField(4, store), VarintField(6, rows), VarintField(7, 1))),
            ArchiveRecord(12, 6002, Message(BytesField(5, Message(VarintField(1, 0),
                BytesField(6, cell), BytesField(7, new byte[] { 0, 0 }))))),
            additionalRecords ?? Message()))), ("preview.png", ValidPreviewPng()));
    }
}
