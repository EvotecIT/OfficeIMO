using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> TrailingCellBoundaryCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote })
            foreach (bool wide in new[] { false, true })
                foreach (bool formula in new[] { false, true })
                    yield return new object[] { kind, wide, formula };
    }

    [Theory]
    [MemberData(nameof(TrailingCellBoundaryCases))]
    public void Out_of_table_offsets_bound_selected_cell_values_and_formula_cache_assessment(
        IWorkDocumentKind kind, bool wide, bool formula) {
        using var package = TableDependencyPackage(kind, Message(), tilePayload:
            BytesField(5, CrossingTrailingCellRow(wide, formula)));
        var source = IWorkSourceDocument.Open(package, kind, new IWorkReadOptions { PreserveSourceRecords = false });
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.True(cell.HasDecodeError);
        Assert.Null(cell.Value);
        Assert.Equal(formula, cell.SourceFormulaIsDeclared);
        foreach (bool visual in new[] { false, true }) {
            package.Position = 0;
            IWorkConversionReport report = ConvertUnitReport(package, kind, visual,
                new IWorkReadOptions { PreserveSourceRecords = false });
            Assert.Empty(report.PreservedRecords);
            IWorkSourceCellIssue issue = Assert.Single(report.SourceCellIssues);
            Assert.Equal((1, 1), (issue.Row, issue.Column));
            Assert.Equal("Truncated cell value field.", issue.Message);
            AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]/7", 1,
                IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
            if (formula) {
                var status = Assert.Single(report.FormulaCells);
                Assert.False(status.ExpressionIsAssessed);
                Assert.Equal(IWorkFormulaCacheStatus.Unassessed, status.CacheStatus);
                Assert.Equal(1, report.FormulaSummary.UnassessedCacheCount);
            } else Assert.Empty(report.FormulaCells);
            Assert.Contains(report.FidelityDiagnostics, d => d.Code == "IWORK_SOURCE_CELLS_UNDECODED"
                && d.LossKind == OfficeConversionLossKind.Unassessed);
        }
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Partial_saved_tables_keep_decode_markers_instead_of_borrowed_record_bytes(IWorkDocumentKind kind) {
        using var package = TableDependencyPackage(kind, Message(), tilePayload:
            BytesField(5, CrossingTrailingCellRow(wide: true, formula: true)));
        var source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        const string marker = "Truncated cell value field.";
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false }); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = WordDocument.Load(saved);
            Assert.Equal(marker, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false }); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = ExcelDocument.Load(saved);
            Assert.Null(reopened.Sheets[0].GetFormulaText(1, 1));
            Assert.Equal(marker, reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
            Assert.Empty(reopened.ValidateOpenXml());
        } else {
            using var automatic = source.ToPowerPointPresentationResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false }); Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options); partial.Value.Save(saved); saved.Position = 0;
            using var reopened = PowerPointPresentation.Load(saved);
            Assert.Equal(marker, Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Fact]
    public void Supported_wide_table_offsets_are_not_limited_by_the_tile_row_stride() {
        byte[] offsets = Enumerable.Repeat((byte)255, 600).ToArray(); offsets[0] = offsets[1] = 0;
        byte[] cell = new byte[20]; cell[0] = 5; cell[1] = 2; WriteUInt32(cell, 8, 2);
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        byte[] row = Message(VarintField(1, 0), BytesField(6, cell), BytesField(7, offsets));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), columns: 300,
            tilePayload: BytesField(5, row));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.Equal(42d, result.Value.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Empty(result.Report.SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Excess_offset_envelopes_remain_unmaterialized_and_recover_readable_rows(IWorkDocumentKind kind) {
        byte[] offsets = new byte[514]; offsets[2] = 12;
        byte[] header = new byte[12]; header[0] = 5; header[1] = 2; WriteUInt32(header, 8, 2 | (1u << 9));
        byte[] row = Message(VarintField(1, 0), BytesField(6, Message(header, new byte[24])), BytesField(7, offsets));
        using var package = TableDependencyPackage(kind, Message(), rows: 2, tilePayload:
            Message(BytesField(5, row), BytesField(5, TileTestRow(1))));
        var cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        Assert.Equal((2, 1, 42d), (cell.Row, cell.Column, Assert.IsType<double>(cell.Value)));
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, visual: false);
        Assert.Empty(report.SourceCellIssues); Assert.Empty(report.FormulaCells);
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]/7", 1,
            IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata);
    }

    [Fact]
    public void Selected_offset_slots_charge_the_existing_source_wide_dimension_budget() {
        // Three slots include two empty trailing declarations; none can hide scanning work.
        byte[] cell = new byte[12]; cell[0] = 5;
        byte[] row = Message(VarintField(1, 0), BytesField(6, cell), BytesField(7, new byte[] { 0, 0, 255, 255, 255, 255 }));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), tilePayload: BytesField(5, row));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTableDimensionEntries = 2 });
        Assert.Contains("dimension", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void Invalid_later_offsets_do_not_retire_a_readable_preceding_cell(bool wide, bool selected) {
        byte[] cell = new byte[20]; cell[0] = 5; cell[1] = 2; WriteUInt32(cell, 8, 2);
        Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        byte[] row = Message(VarintField(1, 0), BytesField(6, cell),
            BytesField(7, new byte[] { 0, 0, 254, 255 }), VarintField(8, wide ? 1ul : 0ul));
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), columns: selected ? 2ul : 1ul,
            tilePayload: BytesField(5, row));
        var table = ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1;
        Assert.Equal(42d, Assert.Single(table.Cells, c => c.Column == 1).Value);
        if (selected) Assert.True(Assert.Single(table.Cells, c => c.Column == 2).HasDecodeError);
        else Assert.Single(table.Cells);
    }

    [Fact]
    public void Row_offset_budget_is_shared_across_selected_rows() {
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), rows: 2,
            tilePayload: Message(BytesField(5, TileTestRow(0)), BytesField(5, TileTestRow(1))));
        var source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTableDimensionEntries = 1 });
        Assert.Contains("dimension", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message,
            StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] CrossingTrailingCellRow(bool wide, bool formula) {
        byte[] header = new byte[12]; header[0] = 5; header[1] = 2;
        WriteUInt32(header, 8, 2 | (formula ? 1u << 9 : 0));
        byte[] other = new byte[24]; other[0] = 5; other[1] = 2;
        WriteUInt32(other, 8, 2); Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, other, 12, 8);
        return Message(VarintField(1, 0), BytesField(6, Message(header, other)),
            BytesField(7, new byte[] { 0, 0, wide ? (byte)3 : (byte)12, 0 }), VarintField(8, wide ? 1ul : 0ul));
    }
}
