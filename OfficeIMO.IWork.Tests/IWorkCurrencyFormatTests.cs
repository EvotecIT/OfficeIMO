using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(2)]
    [InlineData(10)]
    public void Currency_selection_preserves_formula_cache_and_ignores_the_inactive_numeric_slot(int sourceType) {
        byte[] cell = new byte[32]; cell[0] = 5; cell[1] = (byte)sourceType;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 9) | (1u << 13) | (1u << 14));
        Buffer.BlockCopy(BitConverter.GetBytes(-0.5d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, 0); WriteUInt32(cell, 24, 99); WriteUInt32(cell, 28, 1);
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers,
            Message(ReferenceField(22, 13), ReferenceField(6, 14)), cellPayload: cell,
            additionalRecords: Message(
                ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(CurrencyFormat("GBP", 2, negative: 3))))),
                ArchiveRecord(14, 6201, Message(BytesField(3, Message(VarintField(1, 0), BytesField(5, FormulaConstant(1d))))))));
        IWorkTableCell value = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(IWorkCellKind.Formula, value.Kind);
        Assert.Equal(IWorkCellKind.Number, value.ValueKind);
        Assert.Equal(-0.5d, value.Value);
        Assert.True(value.FormulaIsComplete);
        Assert.True(value.CachedValueIsComplete);
        Assert.Equal("GBP", value.NumberFormat!.CurrencyCode);
        package.Position = 0;
        using var converted = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.DoesNotContain(converted.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);
        var cache = reopened.Sheets[0].CellAt(1, 1).GetValue();
        Assert.Equal(global::OfficeIMO.Excel.ExcelCellDataKind.Formula, cache.Kind);
        Assert.Equal(-0.5d, cache.Value);
        Assert.Equal("GBP (0.50)", Assert.Single(reopened.Sheets[0].Range("A1:A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Selected_currency_metadata_is_shared_without_interpreting_a_symbol(IWorkDocumentKind kind) {
        using MemoryStream package = NumberFormatPackage(kind, CurrencyFormat("PLN", 3, accounting: 1), currency: true);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), kind).Item1.Cells);
        Assert.Equal(0.5d, cell.Value);
        Assert.Equal("0.5", cell.DisplayText);
        Assert.Equal(IWorkNumberFormatKind.Currency, cell.NumberFormat!.Kind);
        Assert.Equal("PLN", cell.NumberFormat.CurrencyCode);
        Assert.True(cell.NumberFormat.UseAccountingStyle);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual: false);
        Assert.DoesNotContain(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(report.Diagnostics, d => d.Code == (kind == IWorkDocumentKind.Numbers
            ? "IWORK_NUMBERS_CURRENCY_DISPLAY_APPROXIMATED" : kind == IWorkDocumentKind.Pages
            ? "IWORK_PAGES_NUMBER_FORMAT_APPROXIMATED" : "IWORK_KEYNOTE_NUMBER_FORMAT_APPROXIMATED"));
    }

    [Theory]
    [InlineData(0)] // Missing identifier.
    [InlineData(1)] // Lowercase identifier.
    [InlineData(2)] // Non-letter and unsafe destination format characters.
    [InlineData(3)] // Duplicate identifier.
    [InlineData(4)] // Invalid accounting Boolean.
    [InlineData(5)] // Conflicting accounting/negative-style metadata.
    [InlineData(6)] // Currency selected through the numeric slot.
    public void Unsupported_currency_metadata_keeps_raw_values_and_declaration_evidence(int failure) {
        byte[] format = failure switch {
            0 => Message(NumericFormat(257, 2), VarintField(6, 0)),
            1 => CurrencyFormat("usd", 2),
            2 => CurrencyFormat("U;D", 2),
            3 => Message(CurrencyFormat("USD", 2), BytesField(3, System.Text.Encoding.UTF8.GetBytes("EUR"))),
            4 => CurrencyFormat("USD", 2, accounting: 2),
            5 => CurrencyFormat("USD", 2, negative: 1, accounting: 1),
            _ => CurrencyFormat("USD", 2)
        };
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, format, currency: failure != 6);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(0.5d, cell.Value);
        Assert.Null(cell.NumberFormat);
        Assert.False(cell.HasDecodeError);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Equal("3[1]/6", Assert.Single(report.SourceDeclarationIssues).FieldPath);
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
    }

    [Fact]
    public void Currency_identifiers_consume_the_projection_text_budget() {
        var options = new IWorkReadOptions { MaximumProjectedTextCharacters = 2 };
        using MemoryStream numeric = TableDependencyPackage(IWorkDocumentKind.Numbers, Message(), sheetName: "");
        Assert.NotNull(IWorkSourceDocument.Open(numeric, options).ReadNumbers());
        using MemoryStream currency = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(22, 13),
            cellPayload: CurrencyCell(1), sheetName: "",
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(CurrencyFormat("USD", 2))))));
        Assert.Contains("Text character", Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(currency, options).ReadNumbers()).Message);
    }

    [Fact]
    public void Currency_item_budget_exhaustion_is_fatal_instead_of_partial_format_recovery() {
        byte[] rows = Message(Enumerable.Range(0, 2).Select(row => BytesField(5, Message(VarintField(1, (ulong)row),
            BytesField(6, CurrencyCell((uint)row + 1)), BytesField(7, new byte[] { 0, 0 })))).ToArray());
        using MemoryStream package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(22, 13), rows: 2,
            tilePayload: rows, additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2),
                BytesField(3, FormatEntry(CurrencyFormat("USD", 2))),
                BytesField(3, Message(VarintField(1, 2), BytesField(6, CurrencyFormat("EUR", 2)))))));
        Assert.Contains("Text item", Assert.Throws<InvalidDataException>(() =>
            IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProjectedTextItems = 1 }).ReadNumbers()).Message);
    }

    private static byte[] CurrencyCell(uint identifier) {
        byte[] cell = new byte[24]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 14));
        Buffer.BlockCopy(BitConverter.GetBytes(0.5d), 0, cell, 12, 8);
        WriteUInt32(cell, 20, identifier);
        return cell;
    }

    private static byte[] CurrencyFormat(string code, uint decimals, uint negative = 0, uint grouping = 0, uint accounting = 0) =>
        Message(NumericFormat(257, decimals, negative, grouping), BytesField(3, System.Text.Encoding.UTF8.GetBytes(code)),
            VarintField(6, accounting));
}
