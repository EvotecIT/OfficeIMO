using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Scientific_precision_is_shared_metadata_without_changing_values(IWorkDocumentKind kind) {
        using MemoryStream package = NumberFormatPackage(kind, NumericFormat(259, 4), value: 1250d);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells);
        Assert.Equal(1250d, cell.Value);
        Assert.Equal("1250", cell.DisplayText);
        Assert.Equal(IWorkNumberFormatKind.Scientific, cell.NumberFormat!.Kind);
        Assert.Equal(4, cell.NumberFormat.DecimalPlaces);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.DoesNotContain(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        if (kind != IWorkDocumentKind.Numbers) Assert.Contains(report.Diagnostics, d => d.Code ==
            (kind == IWorkDocumentKind.Pages ? "IWORK_PAGES_NUMBER_FORMAT_OMITTED" : "IWORK_KEYNOTE_NUMBER_FORMAT_OMITTED"));
    }

    [Theory]
    [InlineData(1u, 0u)]
    [InlineData(0u, 1u)]
    public void Unqualified_scientific_negative_styles_and_grouping_retain_source_evidence(uint negative, uint grouping) {
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, NumericFormat(259, 2, negative, grouping), value: -1250d);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(-1250d, cell.Value);
        Assert.Null(cell.NumberFormat);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal("3[1]/6", Assert.Single(report.SourceDeclarationIssues).FieldPath);
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    public void Scientific_selection_preserves_complete_formula_expression_and_numeric_cache(int sourceType) {
        byte[] cell = new byte[28]; cell[0] = 5; cell[1] = (byte)sourceType;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 9) | (1u << 13));
        Buffer.BlockCopy(BitConverter.GetBytes(-0.0625d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, 0); WriteUInt32(cell, 24, 1);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(22, 13), ReferenceField(6, 14)), cellPayload: cell,
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(NumericFormat(259, 3))))),
                ArchiveRecord(14, 6201, Message(BytesField(3, Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d))))))));
        package.Position = 0;
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(converted.IsVisualFallback);
        IWorkTableCell projected = Assert.Single(Assert.Single(Assert.Single(converted.Projection.Sheets).Tables).Cells);
        Assert.Equal(IWorkCellKind.Formula, projected.Kind);
        Assert.True(projected.FormulaIsComplete);
        Assert.True(projected.CachedValueIsComplete);
        Assert.Equal(IWorkNumberFormatKind.Scientific, projected.NumberFormat!.Kind);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);
        var cache = reopened.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(global::OfficeIMO.Excel.ExcelCellDataKind.Formula, cache.Kind);
        Assert.Equal(-0.0625d, cache.Value);
        Assert.Equal("-6.250E-02", Assert.Single(reopened.Sheets[0].Range("A1:A1").CreateVisualSnapshot().Cells).Text);
    }
}
