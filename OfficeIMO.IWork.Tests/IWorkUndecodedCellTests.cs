using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Undecoded_nonformula_cell_has_explicit_unassessed_fidelity(IWorkDocumentKind kind, bool visual) {
        byte[] cell = UndecodedFormulaCell("truncatedValue");
        cell[0] = 4; // Unknown storage cannot establish formula presence.
        using MemoryStream package = TableDependencyPackage(kind, Message(), cellPayload: cell);
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.FormulaCells);
        Assert.Empty(report.PreservedRecords);
        IWorkSourceCellIssue issue = Assert.Single(report.SourceCellIssues);
        Assert.Equal(10ul, issue.TableIdentity!.RecordIdentifier);
        Assert.Equal((1, 1), (issue.Row, issue.Column));
        Assert.Contains("version 4", issue.Message);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkSourceCellIssue>)report.SourceCellIssues).Clear());
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            issue => issue.Code == "IWORK_SOURCE_CELLS_UNDECODED").LossKind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Recovered_native_error_markers_are_not_cell_decode_failures(IWorkDocumentKind kind, bool formula) {
        byte[] cell = new byte[formula ? 16 : 12]; cell[0] = 5; cell[1] = 8;
        WriteUInt32(cell, 8, formula ? 1u << 9 : 0);
        using MemoryStream package = TableDependencyPackage(kind, Message(), cellPayload: cell);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        IWorkTableCell projected = Assert.Single(table.Cells);
        Assert.False(projected.HasDecodeError);
        Assert.Equal("#ERROR", projected.CachedDisplayText);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.Empty(report.SourceCellIssues);
        Assert.DoesNotContain(report.FidelityDiagnostics, item => item.Code == "IWORK_SOURCE_CELLS_UNDECODED"
            || item.Code == "IWORK_TABLE_CELL_DECODE");
        if (formula) Assert.Equal(IWorkFormulaCacheStatus.Approximate, Assert.Single(report.FormulaCells).CacheStatus);
        else Assert.Empty(report.FormulaCells);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Cell_evidence_uses_sparse_coordinates_and_excludes_unmaterialized_storage(IWorkDocumentKind kind) {
        byte[] truncated = UndecodedFormulaCell("truncatedValue").Take(11).ToArray();
        byte[] row = Message(VarintField(1, 1), BytesField(6, truncated),
            BytesField(7, new byte[] { 255, 255, 255, 255, 0, 0 }));
        byte[] tile = Message(BytesField(5, new byte[] { 0x80 }), BytesField(5, row));
        using MemoryStream package = TableDependencyPackage(kind, Message(), rows: 2, columns: 3,
            tilePayload: tile, additionalRecords: ArchiveRecord(99, 6002, tile));
        var options = new IWorkReadOptions { MaximumMaterializedCells = 1 };
        IWorkConversionReport report = ConvertUnitReport(package, kind, readOptions: options);
        IWorkSourceCellIssue issue = Assert.Single(report.SourceCellIssues);
        Assert.Equal((2, 3), (issue.Row, issue.Column));
        Assert.Equal("Truncated cell record.", issue.Message);
        Assert.Empty(report.FormulaCells);
        Assert.Single(report.SourceDeclarationIssues); // Unreadable row has no identified cells.
        using MemoryStream overBudget = TableDependencyPackage(kind, Message(), rows: 2, columns: 3,
            tilePayload: Message(BytesField(5, Message(VarintField(1, 0), BytesField(6, truncated),
                BytesField(7, new byte[] { 0, 0 }))), BytesField(5, row)));
        Assert.Contains("source-wide limit", Assert.Throws<InvalidDataException>(() =>
            ConvertUnitReport(overBudget, kind, readOptions: options)).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Partial_saved_output_retains_existing_decode_marker_and_report_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, Message(),
            cellPayload: UndecodedFormulaCell("truncatedValue"));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        IWorkConversionReport report;
        string marker;
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options); report = partial.Report;
            partial.Value.Save(saved); saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            marker = reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text;
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options); report = partial.Report;
            partial.Value.Save(saved); saved.Position = 0;
            using ExcelDocument reopened = ExcelDocument.Load(saved);
            marker = Assert.IsType<string>(reopened.Sheets[0].CellAt(1, 1).GetValue<string>());
        } else {
            using var automatic = source.ToPowerPointPresentationResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options); report = partial.Report;
            partial.Value.Save(saved); saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            marker = Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text;
        }
        Assert.True(report.IsPartialEditableReconstruction);
        Assert.Equal(Assert.Single(report.SourceCellIssues).Message, marker);
        Assert.Equal("Truncated cell value field.", marker);
    }
}
