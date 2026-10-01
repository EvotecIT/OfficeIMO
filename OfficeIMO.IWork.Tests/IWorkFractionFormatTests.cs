using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Fraction_precision_is_shared_metadata_without_changing_raw_values(IWorkDocumentKind kind) {
        using MemoryStream package = NumberFormatPackage(kind, FractionFormat(16), value: -1.375d);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells);
        Assert.Equal(-1.375d, cell.Value);
        Assert.Equal("-1.375", cell.DisplayText);
        Assert.Equal(IWorkNumberFormatKind.Fraction, cell.NumberFormat!.Kind);
        Assert.Equal(IWorkFractionAccuracy.Sixteenths, cell.NumberFormat.FractionAccuracy);
        Assert.Null(cell.NumberFormat.DecimalPlaces);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.DoesNotContain(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(report.Diagnostics, d => d.Code == (kind == IWorkDocumentKind.Numbers
            ? "IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED" : kind == IWorkDocumentKind.Pages
            ? "IWORK_PAGES_NUMBER_FORMAT_APPROXIMATED" : "IWORK_KEYNOTE_NUMBER_FORMAT_APPROXIMATED"));
    }

    [Theory]
    [InlineData(0)] // Missing precision.
    [InlineData(1)] // Unsupported denominator.
    [InlineData(2)] // Duplicate precision.
    [InlineData(3)] // Wrong precision wire kind.
    [InlineData(4)] // Decimal controls are not qualified for fractions.
    [InlineData(5)] // Nondefault negative style.
    [InlineData(6)] // Grouping.
    [InlineData(7)] // Fraction metadata selected through the currency slot.
    public void Unsupported_fraction_selection_keeps_values_and_physical_declaration_evidence(int failure) {
        byte[] format = failure switch {
            0 => VarintField(1, 262),
            1 => FractionFormat(3),
            2 => Message(FractionFormat(8), VarintField(11, 16)),
            3 => Message(VarintField(1, 262), BytesField(11, Message())),
            4 => Message(FractionFormat(8), VarintField(2, 0)),
            5 => Message(FractionFormat(8), VarintField(4, 1)),
            6 => Message(FractionFormat(8), VarintField(5, 1)),
            _ => FractionFormat(8)
        };
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, format, value: -1.375d, currency: failure == 7);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(-1.375d, cell.Value);
        Assert.Null(cell.NumberFormat);
        Assert.False(cell.HasDecodeError);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal("3[1]/6", Assert.Single(report.SourceDeclarationIssues).FieldPath);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    public void Fraction_selection_preserves_formula_expression_and_numeric_cache(int sourceType) {
        byte[] cell = new byte[28]; cell[0] = 5; cell[1] = (byte)sourceType;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 9) | (1u << 13));
        Buffer.BlockCopy(BitConverter.GetBytes(-1.375d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, 0); WriteUInt32(cell, 24, 1);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(22, 13), ReferenceField(6, 14)), cellPayload: cell,
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(FractionFormat(8))))),
                ArchiveRecord(14, 6201, Message(BytesField(3, Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d))))))));
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(converted.IsVisualFallback);
        IWorkTableCell projected = Assert.Single(Assert.Single(Assert.Single(converted.Projection.Sheets).Tables).Cells);
        Assert.Equal(IWorkCellKind.Formula, projected.Kind);
        Assert.True(projected.FormulaIsComplete);
        Assert.True(projected.CachedValueIsComplete);
        Assert.Equal(IWorkFractionAccuracy.Eighths, projected.NumberFormat!.FractionAccuracy);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);
        var cache = reopened.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(global::OfficeIMO.Excel.ExcelCellDataKind.Formula, cache.Kind);
        Assert.Equal(-1.375d, cache.Value);
        Assert.Equal("-1 3/8", Assert.Single(reopened.Sheets[0].Range("A1:A1").CreateVisualSnapshot().Cells).Text);
    }

    private static byte[] FractionFormat(uint accuracy) => Message(VarintField(1, 262), VarintField(11, accuracy));
}
