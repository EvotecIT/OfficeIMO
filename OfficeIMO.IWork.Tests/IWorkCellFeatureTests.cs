using OfficeIMO.IWork;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_unsupported_cell_features_preserve_values_but_require_explicit_partial_conversion(IWorkDocumentKind kind) {
        using MemoryStream package = TableDependencyPackage(kind, Message(), cellPayload: FeatureCell(empty: false));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.Equal(42d, cell.Value);
        Assert.False(cell.HasDecodeError);
        Assert.Equal(IWorkCellUnsupportedFeatures.ConditionalStyle | IWorkCellUnsupportedFeatures.AppliedConditionalRule
            | IWorkCellUnsupportedFeatures.Comment, cell.UnsupportedFeatures);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.True(report.IsPartialEditableReconstruction);
        AssertTileDeclaration(Assert.Single(report.SourceDeclarationIssues), "5[1]/6", 1,
            IWorkSourceDeclarationIssueKind.UnsupportedField);
        Assert.Empty(report.SourceCellIssues);
        Assert.Contains(report.FidelityDiagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED"
            && diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Feature_only_empty_cells_are_materialized_and_charged_to_the_existing_cell_budget(IWorkDocumentKind kind) {
        byte[] cell = FeatureCell(empty: true);
        using MemoryStream package = TableDependencyPackage(kind, Message(), cellPayload: cell);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.Equal(IWorkCellKind.Empty, Assert.Single(table.Cells).Kind);
        using MemoryStream overBudget = TableDependencyPackage(kind, Message(), columns: 2,
            tilePayload: BytesField(5, Message(VarintField(1, 0), BytesField(6, Message(cell, cell)),
                BytesField(7, new byte[] { 0, 0, (byte)cell.Length, 0 }))));
        Assert.Contains("source-wide limit", Assert.Throws<InvalidDataException>(() =>
            ReadSelectedRichTable(IWorkSourceDocument.Open(overBudget, kind,
                new IWorkReadOptions { MaximumMaterializedCells = 1 }), kind)).Message);
    }

    [Theory]
    [InlineData(7, IWorkCellUnsupportedFeatures.ConditionalStyle)]
    [InlineData(8, IWorkCellUnsupportedFeatures.AppliedConditionalRule)]
    [InlineData(19, IWorkCellUnsupportedFeatures.Comment)]
    public void Individual_selected_feature_flags_remain_distinct(int bit, IWorkCellUnsupportedFeatures expected) {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            cellPayload: FeatureCell(empty: false, 1u << bit));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(expected, cell.UnsupportedFeatures);
    }

    [Fact]
    public void Truncated_feature_fields_keep_decode_evidence_without_assessed_feature_presence() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(),
            cellPayload: FeatureCell(empty: false).Take(24).ToArray());
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.True(cell.HasDecodeError);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers);
        Assert.Single(report.SourceCellIssues);
        Assert.DoesNotContain(report.SourceDeclarationIssues, issue => issue.Kind == IWorkSourceDeclarationIssueKind.UnsupportedField);
    }

    [Fact]
    public void Unused_feature_catalogs_are_not_traversed_or_reported_as_selected_content() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(BytesField(18, new byte[] { 0x80 }), ReferenceField(19, 999)));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceReferenceIssues);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells).UnsupportedFeatures);
    }

    [Fact]
    public void Feature_evidence_deduplicates_shared_rows_and_charges_distinct_rows_to_the_existing_budget() {
        byte[] cell = FeatureCell(empty: false);
        using MemoryStream shared = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), cellPayload: cell, repeatModel: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(shared,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        AssertTileDeclaration(Assert.Single(projection.SourceDeclarationIssues), "5[1]/6", 1,
            IWorkSourceDeclarationIssueKind.UnsupportedField);
        byte[] row(ulong index) => Message(VarintField(1, index), BytesField(6, cell), BytesField(7, new byte[] { 0, 0 }));
        using MemoryStream distinct = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), rows: 2,
            tilePayload: Message(BytesField(5, row(0)), BytesField(5, row(1))));
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(distinct, new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers()).Message);
    }

    [Fact]
    public void Partial_feature_conversion_saves_values_and_reader_retains_the_unassessed_warning() {
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), cellPayload: FeatureCell(empty: false));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        using var automatic = source.ToExcelDocumentResult();
        Assert.True(automatic.IsVisualFallback);
        using var partial = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream(); partial.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        package.Position = 0;
        var reader = new OfficeIMO.Reader.OfficeDocumentReaderBuilder().AddIWorkHandler().Build();
        var result = reader.ReadDocument(package, "features.numbers");
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
        Assert.Equal("42", Assert.Single(Assert.Single(result.Tables).Rows[0]));
    }

    [Fact]
    public void Feature_presence_survives_cross_table_formula_binding_and_saved_formula_output() {
        using MemoryStream package = CrossBindingPackage(sourceCellPayload:
            FeatureCell(empty: false, (1u << 9) | (1u << 19)));
        using var result = IWorkSourceDocument.Open(package).ToExcelDocumentResult(new IWorkConversionOptions {
            AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true
        });
        IWorkTableCell cell = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.True(cell.FormulaIsComplete);
        Assert.Equal(IWorkCellUnsupportedFeatures.Comment, cell.UnsupportedFeatures);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.StartsWith("SUM(", reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    private static byte[] FeatureCell(bool empty, uint featureFlags = (1u << 7) | (1u << 8) | (1u << 19)) {
        uint flags = featureFlags | (empty ? 0u : 1u << 1);
        int featureCount = Enumerable.Range(0, 21).Count(bit => (featureFlags & (1u << bit)) != 0);
        byte[] cell = new byte[12 + (empty ? 0 : 8) + featureCount * 4];
        cell[0] = 5;
        cell[1] = empty ? (byte)0 : (byte)2;
        WriteUInt32(cell, 8, flags);
        if (!empty) Buffer.BlockCopy(BitConverter.GetBytes(42d), 0, cell, 12, 8);
        return cell;
    }
}
