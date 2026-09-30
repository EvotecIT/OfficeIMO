using OfficeIMO.IWork;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Malformed_selected_tile_payload_retains_evidence_and_allows_explicit_partial_or_preview(
        IWorkDocumentKind kind, bool visual) {
        using MemoryStream package = TableDependencyPackage(kind, Message(), tilePayload: new byte[] { 0x80 });
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        AssertTileDeclaration(issue, "$", null, IWorkSourceDeclarationIssueKind.MalformedMessage);
        Assert.Empty(report.SourceReferenceIssues);
        Assert.Equal(visual ? IWorkProjectionKind.VisualFallback : IWorkProjectionKind.EditableReconstruction,
            report.ProjectionKind);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED").LossKind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Unreadable_physical_tile_row_keeps_its_position_and_recovers_readable_sibling(
        IWorkDocumentKind kind, bool wrongWire) {
        byte[] rows = Message(wrongWire ? VarintField(5, 999) : BytesField(5, new byte[] { 0x80 }),
            BytesField(5, TileTestRow(1)));
        using MemoryStream package = TableDependencyPackage(kind, Message(), rows: 2, tilePayload: rows);
        (IWorkTable table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
        IWorkTableCell cell = Assert.Single(table.Cells);
        Assert.Equal(2, cell.Row);
        Assert.Equal(42d, cell.Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]", 1,
            IWorkSourceDeclarationIssueKind.MalformedMessage);
        Assert.DoesNotContain(report.SourceUnits, unit => unit.Identity.RecordIdentifier == 999);
        Assert.Empty(report.FormulaCells);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Recovered_tile_row_survives_partial_destination_save_and_reopen(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, Message(), rows: 2,
            tilePayload: Message(BytesField(5, new byte[] { 0x80 }), BytesField(5, TileTestRow(1))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            Assert.Equal("42", reopened.Tables[0].Rows[1].Cells[0].Paragraphs[0].Text);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(saved);
            Assert.Equal(42d, reopened.Sheets[0].CellAt(2, 1).GetValue<double>());
        } else {
            using var automatic = source.ToPowerPointPresentationResult();
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            Assert.Equal("42", reopened.Slides[0].Tables.First().GetCell(1, 0).Text);
        }
    }

    private static byte[] TileTestRow(ulong row) {
        byte[] cell = new byte[20]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, 1u << 1);
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        return Message(VarintField(1, row), BytesField(6, cell), BytesField(7, new byte[] { 0, 0 }));
    }

    public static IEnumerable<object[]> InvalidTileRows() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (string failure in new[] { "missingOffsets", "oddOffsets", "wideFlag", "duplicateBuffer",
                "missingIndex", "indexRange", "tableRange", "legacy", "duplicateOffsets", "trailingOffset" })
                yield return new object[] { kind, failure };
    }

    [Theory]
    [MemberData(nameof(InvalidTileRows))]
    public void Rejected_row_selection_and_storage_retains_bounded_physical_evidence(IWorkDocumentKind kind, string failure) {
        byte[] row = InvalidTileTestRow(failure);
        using MemoryStream package = TableDependencyPackage(kind, Message(),
            columns: failure == "duplicateOffsets" ? 2ul : 1ul, tilePayload: BytesField(5, row));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 });
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]", 1,
            failure == "legacy" ? IWorkSourceDeclarationIssueKind.RejectedMessageSet
                : IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Empty(report.FormulaCells);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Invalid_storage_after_unreadable_row_retains_the_physical_index(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, Message(), rows: 3,
            tilePayload: Message(BytesField(5, new byte[] { 0x80 }), BytesField(5, TileTestRow(1)),
                BytesField(5, InvalidTileTestRow("oddOffsets"))));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.Equal(new[] { "5[1]", "5[3]" }, report.SourceDeclarationIssues.Select(issue => issue.FieldPath));
        Assert.All(report.SourceDeclarationIssues, issue => Assert.Equal(12ul, issue.Owner.RecordIdentifier));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Tile_row_evidence_budget_is_cumulative_and_fatal(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, Message(), rows: 2,
            tilePayload: Message(BytesField(5, new byte[] { 0x80 }), BytesField(5, new byte[] { 0x80 })));
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind)).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Configured_tile_or_nested_row_field_limit_is_not_recovered(IWorkDocumentKind kind, bool nested) {
        byte[] fields = Message(Enumerable.Range(0, 9).Select(_ => VarintField(1, 0)).ToArray());
        using MemoryStream package = TableDependencyPackage(kind, Message(),
            tilePayload: nested ? BytesField(5, fields) : fields);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => ReadSelectedRichTable(source, kind)).Message);
    }

    [Fact]
    public void Reused_tile_evidence_reports_each_physical_path_once() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            repeatModel: true, tilePayload: BytesField(5, InvalidTileTestRow("wideFlag")));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        AssertTileDeclaration(Assert.Single(projection.SourceDeclarationIssues), "5[1]", 1,
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Unsupported_tile_envelopes_retain_whole_payload_evidence(bool excessiveMetadata) {
        byte[] payload = excessiveMetadata
            ? Message(Enumerable.Range(0, 8).Select(_ => VarintField(1, 0)).ToArray())
            : VarintField(99, 0);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), tilePayload: payload);
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "$", null,
            IWorkSourceDeclarationIssueKind.RejectedMessageSet);
    }

    [Fact]
    public void Rejected_tile_row_count_is_reported_before_nested_content_is_parsed() {
        byte[] overLimitRow = Message(Enumerable.Range(0, 9).Select(_ => VarintField(1, 0)).ToArray());
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            tilePayload: Message(BytesField(5, overLimitRow), BytesField(5, new byte[] { 0x80 })));
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers,
            readOptions: new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5", 2,
            IWorkSourceDeclarationIssueKind.RejectedMessageSet);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_TILE_ROWS_UNSUPPORTED");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Invalid_tile_index_records_its_entry_without_selecting_the_target(bool missingIndex) {
        byte[] entry = Message(missingIndex ? Message() : VarintField(1, 1), ReferenceField(2, 999));
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            BytesField(3, BytesField(1, entry)), includeTile: false);
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal("4/3/1[1]", issue.FieldPath);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
        Assert.Empty(report.SourceReferenceIssues);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Duplicate_tile_index_or_identity_retains_the_second_physical_entry(bool duplicateIdentity) {
        byte[] entries = Message(BytesField(1, Message(VarintField(1, 0), ReferenceField(2, 12))),
            BytesField(1, Message(VarintField(1, duplicateIdentity ? 1ul : 0ul),
                ReferenceField(2, duplicateIdentity ? 12ul : 999ul))));
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            BytesField(3, entries), includeTile: false, rows: 257);
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(11ul, issue.Owner.RecordIdentifier);
        Assert.Equal("4/3/1[2]", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Empty(report.SourceReferenceIssues);
    }

    [Fact]
    public void Duplicate_row_index_retains_the_first_cell_and_reports_the_later_entry() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), rows: 2,
            tilePayload: Message(BytesField(5, TileTestRow(0)), BytesField(5, TileTestRow(0))));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
        AssertTileDeclaration(Assert.Single(projection.SourceDeclarationIssues), "5[2]", 1,
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
    }

    private static byte[] InvalidTileTestRow(string failure) {
        byte[] cell = new byte[20]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, 1u << 1); Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        if (failure == "legacy") return Message(VarintField(1, 0), BytesField(3, new byte[] { 1 }), BytesField(4, new byte[] { 1 }));
        byte[] index = failure == "missingIndex" ? Message()
            : VarintField(1, failure == "indexRange" ? 256ul : failure == "tableRange" ? 1ul : 0ul);
        byte[] offsets = failure == "oddOffsets" ? new byte[] { 0 }
            : failure is "duplicateOffsets" or "trailingOffset" ? new byte[] { 0, 0, 0, 0 } : new byte[] { 0, 0 };
        return Message(index, BytesField(6, cell), failure == "duplicateBuffer" ? BytesField(6, cell) : Message(),
            failure == "missingOffsets" ? Message() : BytesField(7, offsets),
            failure == "wideFlag" ? VarintField(8, 2) : Message());
    }

    private static void AssertTileDeclaration(IWorkSourceDeclarationIssue issue, string path, int? count,
        IWorkSourceDeclarationIssueKind kind) {
        Assert.Equal(12ul, issue.Owner.RecordIdentifier);
        Assert.Equal(path, issue.FieldPath);
        Assert.Equal(count, issue.DeclaredValueCount);
        Assert.Equal(kind, issue.Kind);
    }
}
