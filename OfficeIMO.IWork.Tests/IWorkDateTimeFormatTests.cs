using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> DatePatternCases() {
        yield return new object[] { "dd/MM/y", "dd/mm/yyyy", "05/01/2026" };
        yield return new object[] { "dd/MM/y HH:mm", "dd/mm/yyyy hh:mm", "05/01/2026 13:04" };
        yield return new object[] { "HH:mm:ss", "hh:mm:ss", "13:04:09" };
        yield return new object[] { "h:mm a", "h:mm am/pm", "1:04 pm" };
        yield return new object[] { "d MMM yyyy", "d mmm yyyy", "5 Jan 2026" };
    }

    [Theory]
    [MemberData(nameof(DatePatternCases))]
    public void Qualified_source_date_patterns_preserve_typed_values_and_saved_xlsx(string pattern, string code, string expected) {
        DateTime value = new(2026, 1, 5, 13, 4, 9, DateTimeKind.Utc);
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, pattern, value);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        var cell = result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!;
        Assert.Equal(value, cell.Value);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.NotNull(cell.NumberFormat);
        Assert.Equal("DateTime", cell.NumberFormat.Kind.ToString());
        Assert.Equal(code, cell.NumberFormat.ToSpreadsheetFormatCode());
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DATE_DISPLAY_APPROXIMATED");
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(value.ToOADate(), reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(code, reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal(expected, Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [MemberData(nameof(DatePatternCases))]
    public void Qualified_source_date_patterns_use_shared_text_in_saved_tables_and_Reader(string pattern, string code, string expected) {
        DateTime value = new(2026, 1, 5, 13, 4, 9, DateTimeKind.Utc);
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Keynote }) {
            using var package = DateFormatPackage(kind, pattern, value);
            var source = IWorkSourceDocument.Open(package, kind);
            var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
            Assert.Equal(value, cell.Value);
            Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
            Assert.Equal(code, cell.NumberFormat!.ToSpreadsheetFormatCode());
            Assert.True(cell.TryGetFormattedNumber(out string text, out _)); Assert.Equal(expected, text);
            using var saved = new MemoryStream();
            if (kind == IWorkDocumentKind.Pages) {
                using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
                result.Value.Save(saved); saved.Position = 0;
                using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
                Assert.Equal(expected, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Empty(reopened.ValidateDocument());
            } else {
                using var result = source.ToPowerPointPresentationResult(); Assert.False(result.IsVisualFallback);
                result.Value.Save(saved); saved.Position = 0;
                using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
                Assert.Equal(expected, Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
                Assert.Empty(reopened.ValidateDocument());
            }
        }
        using var numbers = DateFormatPackage(IWorkDocumentKind.Numbers, pattern, value);
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(numbers, "date.numbers");
        Assert.Equal(expected, Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
    }

    [Theory]
    [InlineData(15, "dd/MM/y", "05/01/0015")]
    [InlineData(15, "d MMM yyyy", "5 Jan 0015")]
    public void Qualified_early_year_display_does_not_weaken_xlsx_date_safety(int year, string pattern, string expected) {
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, pattern, new DateTime(year, 1, 5, 0, 0, 0, DateTimeKind.Utc));
        var source = IWorkSourceDocument.Open(package);
        var cell = Assert.Single(source.ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.True(cell.TryGetFormattedNumber(out string text, out _)); Assert.Equal(expected, text);
        using var result = source.ToExcelDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Qualified_date_patterns_do_not_bypass_xlsx_submillisecond_precision_guards() {
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, "HH:mm:ss", new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddTicks(1));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
    }

    [Theory]
    [InlineData(0)] // Unqualified pattern.
    [InlineData(1)] // Suppress date.
    [InlineData(2)] // Suppress time.
    [InlineData(3)] // Ambiguous pattern.
    [InlineData(4)] // Wrong wire kind.
    [InlineData(5)] // Invalid UTF8.
    [InlineData(6)] // Custom controls cannot be ignored.
    [InlineData(7)] // Non-Boolean suppression.
    public void Unqualified_date_formats_keep_raw_dates_and_catalog_evidence(int failure) {
        byte[] format = failure switch {
            0 => DateFormat("yyyy 'era' G"),
            1 => Message(DateFormat("dd/MM/y"), VarintField(12, 1)),
            2 => Message(DateFormat("dd/MM/y"), VarintField(13, 1)),
            3 => Message(DateFormat("dd/MM/y"), StringField(14, "dd/MM/y")),
            4 => Message(VarintField(1, 261), VarintField(14, 1)),
            5 => Message(VarintField(1, 261), BytesField(14, new byte[] { 0xff })),
            6 => Message(DateFormat("dd/MM/y"), VarintField(99, 0)),
            _ => Message(DateFormat("dd/MM/y"), VarintField(12, 2))
        };
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5), format: format);
        var cell = Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.IsType<DateTime>(cell.Value); Assert.Null(cell.NumberFormat);
        Assert.Equal(IWorkCellUnsupportedFeatures.DateFormat, cell.UnsupportedFeatures);
        package.Position = 0; var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        Assert.Contains(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 13 && d.FieldPath == "3[1]/6");
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
    }

    [Fact]
    public void Qualified_date_patterns_preserve_blank_cells_and_consume_existing_budgets() {
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5), empty: true);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.Single(result.Projection.Sheets[0].Tables[0].Cells);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal("dd/mm/yyyy", reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal("", Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
        using var limited = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5), empty: true, columns: 2);
        Assert.Contains("source-wide limit", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(limited,
            new IWorkReadOptions { MaximumMaterializedCells = 1 }).ReadNumbers()).Message);
    }

    private static byte[] DateFormat(string pattern) => Message(VarintField(1, 261), StringField(14, pattern));

    [Fact]
    public void Qualified_date_metadata_is_charged_once_per_selected_format_and_budget_failures_are_fatal() {
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5),
            empty: true, columns: 2, sheetName: "");
        var projection = IWorkSourceDocument.Open(package, new IWorkReadOptions {
            MaximumProjectedTextCharacters = 7, MaximumProjectedTextItems = 1
        }).ReadNumbers();
        Assert.Equal(2, projection.Sheets[0].Tables[0].Cells.Count);
        Assert.Same(projection.Sheets[0].Tables[0].Cells[0].NumberFormat, projection.Sheets[0].Tables[0].Cells[1].NumberFormat);
        package.Position = 0;
        Assert.Contains("Text character", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 6 }).ReadNumbers()).Message);
        using var oversized = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5),
            format: Message(DateFormat("dd/MM/y"), Message(Enumerable.Repeat(VarintField(99, 0), 17).ToArray())));
        Assert.Contains("field", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(oversized,
            new IWorkReadOptions { MaximumProtobufFieldCount = 16 }).ReadNumbers()).Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Explicit_false_date_suppression_flags_preserve_the_qualified_pattern() {
        using var package = DateFormatPackage(IWorkDocumentKind.Numbers, "dd/MM/y", new DateTime(2026, 1, 5),
            format: Message(DateFormat("dd/MM/y"), VarintField(12, 0), VarintField(13, 0)));
        var cell = Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.Equal("dd/MM/y", cell.NumberFormat!.DateTimeFormat!.SourcePattern);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
    }

    [Fact]
    public void Empty_cells_select_date_format_without_activating_dormant_duration() {
        byte[] cell = new byte[24]; cell[0] = 5;
        WriteUInt32(cell, 8, (1u << 12) | (1u << 15) | (1u << 16)); WriteUInt32(cell, 12, 3); WriteUInt32(cell, 16, 1); WriteUInt32(cell, 20, 2);
        using var package = TableDependencyPackage(IWorkDocumentKind.Numbers, ReferenceField(22, 13), cellPayload: cell,
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2),
                BytesField(3, FormatEntry(DateFormat("dd/MM/y"))),
                BytesField(3, Message(VarintField(1, 2), BytesField(6, DurationFormat()))))));
        var actual = Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.Equal(IWorkNumberFormatKind.DateTime, actual.NumberFormat!.Kind); Assert.Null(actual.Value);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, actual.UnsupportedFeatures);
    }

    private static MemoryStream DateFormatPackage(IWorkDocumentKind kind, string pattern, DateTime value,
        byte[]? format = null, bool empty = false, int columns = 1, string sheetName = "Sheet") {
        byte[] cell = new byte[empty ? 20 : 24]; cell[0] = 5; cell[1] = empty ? (byte)0 : (byte)5;
        WriteUInt32(cell, 8, (empty ? 1u << 12 : 1u << 2) | 1u << 15);
        if (empty) WriteUInt32(cell, 12, 3);
        if (!empty) Buffer.BlockCopy(BitConverter.GetBytes((value - new DateTime(2001, 1, 1)).TotalSeconds), 0, cell, 12, 8);
        WriteUInt32(cell, cell.Length - 4, 1);
        byte[] offsets = new byte[columns * 2];
        for (int c = 0; c < columns; c++) offsets[c * 2] = (byte)(c * cell.Length);
        return TableDependencyPackage(kind, ReferenceField(22, 13), columns: (ulong)columns, sheetName: sheetName,
            tilePayload: BytesField(5, Message(VarintField(1, 0), BytesField(6, Message(Enumerable.Repeat(cell, columns).ToArray())), BytesField(7, offsets))),
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(format ?? DateFormat(pattern))))));
    }
}
