using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Qualified_duration_only_empty_cells_preserve_semantic_and_saved_xlsx_formats(IWorkDocumentKind kind) {
        using var package = EmptyDurationFormatPackage(kind);
        var source = IWorkSourceDocument.Open(package, kind);
        var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.Equal(IWorkCellKind.Empty, cell.Kind);
        Assert.Null(cell.Value);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.Equal(IWorkNumberFormatKind.Duration, cell.NumberFormat!.Kind);
        if (kind != IWorkDocumentKind.Numbers) return;
        using var converted = source.ToExcelDocumentResult();
        Assert.False(converted.IsVisualFallback);
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal("[h]\"h\" m\"m\"", reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal("", Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Fact]
    public void Qualified_duration_only_empty_cells_consume_the_materialization_budget() {
        using var package = EmptyDurationFormatPackage(IWorkDocumentKind.Numbers, columns: 2);
        Assert.Contains("source-wide limit", Assert.Throws<InvalidDataException>(() => IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumMaterializedCells = 1 }).ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Qualified_duration_format_retains_source_seconds_and_shared_display(IWorkDocumentKind kind) {
        using var package = DurationFormatPackage(kind, 8640);
        var source = IWorkSourceDocument.Open(package, kind);
        var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.Equal(8640d, cell.Value);
        Assert.Equal("8640s", cell.CachedDisplayText);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.NotNull(cell.NumberFormat);
        Assert.Equal("Duration", cell.NumberFormat.Kind.ToString());
        Assert.Equal("[h]\"h\" m\"m\"", cell.NumberFormat.ToSpreadsheetFormatCode());
        Assert.True(cell.TryGetFormattedNumber(out string text, out _));
        Assert.Equal("2h 24m", text);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, kind, visual: false).SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(0d, "0h 0m")]
    [InlineData(8640d, "2h 24m")]
    [InlineData(-8640d, "-2h 24m")]
    [InlineData(129600d, "36h 0m")]
    public void Qualified_duration_format_survives_saved_xlsx(double seconds, string expected) {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, seconds);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(seconds / 86400, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(expected, Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Qualified_duration_format_reaches_saved_text_tables(IWorkDocumentKind kind) {
        using var package = DurationFormatPackage(kind, -8640);
        var source = IWorkSourceDocument.Open(package, kind);
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.Equal("-2h 24m", reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else {
            using var result = source.ToPowerPointPresentationResult(); Assert.False(result.IsVisualFallback);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            Assert.Equal("-2h 24m", Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Fact]
    public void Reader_duration_display_uses_shared_format_and_reports_approximation() {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 8640);
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "duration.numbers");
        Assert.Equal("2h 24m", Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
    }

    [Fact]
    public void Duration_outside_the_shared_display_range_preserves_raw_text_and_reports_omission() {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 1e308);
        var cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.False(cell.TryGetFormattedNumber(out string raw, out _));
        Assert.Equal(cell.CachedDisplayText, raw);
        package.Position = 0;
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "large.numbers");
        Assert.Equal(raw, Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_OMITTED");
    }

    [Fact]
    public void Reader_duration_formatting_and_diagnostics_respect_projected_header_and_cell_limits() {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 8640);
        var source = IWorkSourceDocument.Open(package);
        var format = source.ReadNumbers().Sheets[0].Tables[0].Cells[0].NumberFormat;
        var table = new IWorkTable("Limited", 3, 2, new[] {
            new IWorkTableCell(1, 1, IWorkCellKind.Duration, 8640d, numberFormat: format),
            new IWorkTableCell(1, 2, IWorkCellKind.Duration, 1e308, numberFormat: format),
            new IWorkTableCell(2, 1, IWorkCellKind.Duration, 1e308, numberFormat: format),
            new IWorkTableCell(3, 1, IWorkCellKind.Duration, 1e308, numberFormat: format)
        }, headerRowCount: 2);
        var numbers = new IWorkNumbersProjection(source, new[] {
            new IWorkNumbersSheet("Sheet", new[] { table }, Array.Empty<string>())
        }, Array.Empty<IWorkDiagnostic>(), supportsEditableReconstruction: true);
        var result = new OfficeDocumentReadResult();
        var projection = new IWorkReadProjection(result, "limited.numbers", new ReaderOptions { MaxTableRows = 1 },
            new ReaderIWorkOptions { MaximumTableColumns = 1, MaximumProjectedTableCells = 1 }, System.Threading.CancellationToken.None);
        projection.AddNumbers(numbers); projection.Complete(source);
        Assert.Equal(new[] { "2h 24m" }, Assert.Single(result.Tables).Columns);
        Assert.Empty(result.Tables[0].Rows);
        Assert.Contains(result.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
        Assert.DoesNotContain(result.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_OMITTED");
    }

    [Theory]
    [InlineData(0)] // Automatic units.
    [InlineData(1)] // Other unit range.
    [InlineData(2)] // Other label style.
    [InlineData(3)] // Duplicate setting.
    [InlineData(4)] // Wrong-wire setting.
    [InlineData(5)] // Extra custom/scaling metadata.
    [InlineData(6)] // Missing setting; do not infer the qualified explicit range.
    public void Unqualified_duration_settings_retain_cache_and_physical_evidence(int failure) {
        byte[] format = failure switch {
            0 => DurationFormat(automatic: 1),
            1 => DurationFormat(largest: 2),
            2 => DurationFormat(style: 0),
            3 => Message(DurationFormat(), VarintField(15, 4)),
            4 => Message(VarintField(1, 268), VarintField(7, 1), BytesField(15, new byte[] { 4 }), VarintField(16, 8), VarintField(40, 0)),
            5 => Message(DurationFormat(), VarintField(19, 0)),
            _ => Message(VarintField(1, 268), VarintField(7, 1), VarintField(15, 4), VarintField(16, 8))
        };
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 8640, format);
        var cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(8640d, cell.Value); Assert.Null(cell.NumberFormat);
        Assert.Equal(IWorkCellUnsupportedFeatures.DurationFormat, cell.UnsupportedFeatures);
        package.Position = 0;
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        Assert.Contains(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 13 && d.FieldPath == "3[1]/6");
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
    }

    private static byte[] DurationFormat(ulong style = 1, ulong largest = 4, ulong smallest = 8, ulong automatic = 0) =>
        Message(VarintField(1, 268), VarintField(7, style), VarintField(15, largest), VarintField(16, smallest), VarintField(40, automatic));

    private static MemoryStream EmptyDurationFormatPackage(IWorkDocumentKind kind, int columns = 1) {
        byte[] cell = new byte[16]; cell[0] = 5;
        WriteUInt32(cell, 8, 1u << 16); WriteUInt32(cell, 12, 1);
        byte[] offsets = new byte[columns * 2];
        for (int column = 0; column < columns; column++) offsets[column * 2] = (byte)(column * cell.Length);
        return TableDependencyPackage(kind, ReferenceField(22, 13), columns: (ulong)columns,
            tilePayload: BytesField(5, Message(VarintField(1, 0),
                BytesField(6, Message(Enumerable.Repeat(cell, columns).ToArray())), BytesField(7, offsets))),
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(DurationFormat())))));
    }

    private static MemoryStream DurationFormatPackage(IWorkDocumentKind kind, double seconds, byte[]? format = null) {
        byte[] cell = new byte[24]; cell[0] = 5; cell[1] = 7;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 16));
        Buffer.BlockCopy(BitConverter.GetBytes(seconds), 0, cell, 12, 8); WriteUInt32(cell, 20, 1);
        return TableDependencyPackage(kind, ReferenceField(22, 13), cellPayload: cell,
            additionalRecords: ArchiveRecord(13, 6005, Message(VarintField(1, 2), BytesField(3, FormatEntry(format ?? DurationFormat())))));
    }
}
