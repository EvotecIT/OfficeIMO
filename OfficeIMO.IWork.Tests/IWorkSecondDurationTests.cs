using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> SecondDurationContainers() {
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote }) {
            yield return new object[] { kind, IWorkDurationUnit.Hour, -129661d, "[h]\"h\" m\"m\" s\"s\"", "-36h 1m 1s" };
            yield return new object[] { kind, IWorkDurationUnit.Minute, -129661d, "[m]\"m\" s\"s\"", "-2161m 1s" };
            yield return new object[] { kind, IWorkDurationUnit.Second, -129661d, "[s]\"s\"", "-129661s" };
            foreach (double seconds in new[] { 11d, -11d }) {
                string sign = seconds < 0 ? "-" : "";
                yield return new object[] { kind, IWorkDurationUnit.Hour, seconds, "[h]\"h\" m\"m\" s\"s\"", sign + "0h 0m 11s" };
                yield return new object[] { kind, IWorkDurationUnit.Minute, seconds, "[m]\"m\" s\"s\"", sign + "0m 11s" };
                yield return new object[] { kind, IWorkDurationUnit.Second, seconds, "[s]\"s\"", sign + "11s" };
            }
        }
    }

    [Theory]
    [MemberData(nameof(SecondDurationContainers))]
    public void Fixed_second_ranges_preserve_raw_seconds_and_editable_saved_table_display(
        IWorkDocumentKind kind, IWorkDurationUnit largest, double seconds, string code, string expected) {
        using var package = DurationFormatPackage(kind, seconds, DurationFormat(largest: (ulong)largest, smallest: 16));
        var source = IWorkSourceDocument.Open(package, kind);
        var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.Equal(seconds, cell.Value);
        Assert.Equal(IWorkCellKind.Duration, cell.ValueKind);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.Equal(largest, cell.NumberFormat!.DurationFormat!.LargestUnit);
        Assert.Equal(IWorkDurationUnit.Second, cell.NumberFormat.DurationFormat.SmallestUnit);
        Assert.Equal(code, cell.NumberFormat.ToSpreadsheetFormatCode());
        Assert.True(cell.TryGetFormattedNumber(out string display, out _));
        Assert.Equal(expected, display);
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult();
            result.Report.RequireCompleteEditableReconstruction();
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED");
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            Assert.Equal(seconds / 86400, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Equal(code, reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
            Assert.Equal(expected, Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
        } else if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.Equal(expected, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else {
            using var result = source.ToPowerPointPresentationResult();
            result.Report.RequireCompleteEditableReconstruction();
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            Assert.Equal(expected, Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData(IWorkDurationUnit.Hour, "[h]\"h\" m\"m\" s\"s\"")]
    [InlineData(IWorkDurationUnit.Minute, "[m]\"m\" s\"s\"")]
    [InlineData(IWorkDurationUnit.Second, "[s]\"s\"")]
    public void Fixed_second_ranges_preserve_explicit_formats_on_blank_saved_cells(IWorkDurationUnit largest, string code) {
        using var package = EmptyDurationFormatPackage(IWorkDocumentKind.Numbers,
            format: DurationFormat(largest: (ulong)largest, smallest: 16));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(IWorkCellKind.Empty, Assert.Single(result.Projection.Sheets[0].Tables[0].Cells).ValueKind);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(code, reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal("", Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [InlineData(IWorkDurationUnit.Hour, "36h 1m 1s")]
    [InlineData(IWorkDurationUnit.Minute, "2161m 1s")]
    [InlineData(IWorkDurationUnit.Second, "129661s")]
    public void Reader_fixed_second_ranges_use_total_elapsed_units_and_report_display_approximation(
        IWorkDurationUnit largest, string expected) {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 129661,
            DurationFormat(largest: (ulong)largest, smallest: 16));
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "seconds.numbers");
        Assert.Equal(expected, Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
    }

    [Theory]
    [InlineData(32UL, 32UL)] // Total milliseconds are a separate native profile.
    [InlineData(8UL, 8UL)] // Minute-only declarations are a separate native profile.
    [InlineData(16UL, 4UL)] // Reversed unit range.
    [InlineData(1UL, 16UL)] // Weeks have no elapsed XLSX unit.
    public void Other_duration_ranges_retain_raw_cache_and_unsupported_format_evidence(ulong largest, ulong smallest) {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, 3661.75,
            DurationFormat(largest: largest, smallest: smallest));
        var cell = Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.Equal(3661.75d, cell.Value);
        Assert.Null(cell.NumberFormat);
        Assert.Equal(IWorkCellUnsupportedFeatures.DurationFormat, cell.UnsupportedFeatures);
        package.Position = 0;
        var report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false);
        Assert.Contains(report.SourceDeclarationIssues, d => d.Owner.RecordIdentifier == 13 && d.FieldPath == "3[1]/6");
        Assert.Contains(report.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
    }
}
