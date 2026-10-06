using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Row_cell_count_mismatch_retains_recoverable_cells_and_native_path(IWorkDocumentKind kind) {
        byte[] row = Message(TileTestRow(0), VarintField(2, 2));
        using MemoryStream package = TableDependencyPackage(kind, Message(), tilePayload: BytesField(5, row));
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.Equal(42d, Assert.Single(table.Cells).Value);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.True(report.IsPartialEditableReconstruction);
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]/2", 1,
            IWorkSourceDeclarationIssueKind.InvalidValue);
        Assert.Empty(report.SourceCellIssues);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_ROW_STORAGE_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData("wrongWire", 1)]
    [InlineData("duplicate", 2)]
    [InlineData("huge", 1)]
    public void Unreadable_or_ambiguous_row_count_is_bounded_without_discarding_values(string failure, int outerCount) {
        byte[] count = failure == "wrongWire" ? BytesField(2, new byte[] { 1 })
            : failure == "duplicate" ? Message(VarintField(2, 1), VarintField(2, 1))
            : VarintField(2, ulong.MaxValue);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            tilePayload: BytesField(5, Message(TileTestRow(0), count)), repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.All(Assert.Single(projection.Sheets).Tables, table => Assert.Equal(42d, Assert.Single(table.Cells).Value));
        AssertTileDeclaration(Assert.Single(projection.SourceDeclarationIssues), "5[1]/2", outerCount,
            IWorkSourceDeclarationIssueKind.InvalidValue);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Unselected_nonempty_modern_buffer_has_evidence_without_invented_cell_coordinates(IWorkDocumentKind kind, bool visual) {
        byte[] row = Message(VarintField(1, 0), BytesField(6, new byte[] { 5, 2, 0 }),
            BytesField(7, new byte[] { 255, 255 }));
        using MemoryStream package = TableDependencyPackage(kind, Message(), tilePayload: BytesField(5, row));
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual,
            new IWorkReadOptions { PreserveSourceRecords = false });
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]/6", 1,
            IWorkSourceDeclarationIssueKind.InvalidValue);
        Assert.Empty(report.SourceCellIssues);
        Assert.Empty(report.FormulaCells);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_ROW_STORAGE_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Valid_row_counts_measure_physical_records_including_unformatted_empty_cells(bool storedEmpty) {
        byte[] emptyCell = storedEmpty ? new byte[12] : Array.Empty<byte>();
        if (storedEmpty) emptyCell[0] = 5;
        byte[] row = Message(VarintField(1, 0), VarintField(2, storedEmpty ? 1ul : 0ul),
            BytesField(6, emptyCell), BytesField(7, storedEmpty ? new byte[] { 0, 0 } : new byte[] { 255, 255 }));
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), tilePayload: BytesField(5, row));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.Empty(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.True(projection.HasEditableContent);
    }

    [Fact]
    public void Row_storage_evidence_uses_shared_declaration_budget_across_physical_paths() {
        byte[] row = Message(VarintField(1, 0), VarintField(2, 1), BytesField(6, new byte[] { 5 }), BytesField(7, new byte[] { 255, 255 }));
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), tilePayload: BytesField(5, row));
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers()).Message);
    }

    [Fact]
    public void Inconsistent_row_count_requires_partial_policy_and_preserves_saved_numeric_value() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            tilePayload: BytesField(5, Message(TileTestRow(0), VarintField(2, 2))));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(automatic.IsVisualFallback);
        using var partial = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        Assert.Throws<InvalidOperationException>(() => partial.Report.RequireCompleteEditableReconstruction());
        using var saved = new MemoryStream();
        partial.Value.Save(saved);
        saved.Position = 0;
        using OfficeIMO.Excel.ExcelDocument reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }
}
