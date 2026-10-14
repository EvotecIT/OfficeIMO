using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> MillisecondDurationContainers() {
        foreach (var kind in new[] { IWorkDocumentKind.Pages, IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote }) {
            yield return new object[] { kind, IWorkDurationUnit.Hour, "[h]\"h\" m\"m\" s.000\"s\" ", "-36h 1m 1.750s " };
            yield return new object[] { kind, IWorkDurationUnit.Minute, "[m]\"m\" s.000\"s\" ", "-2161m 1.750s " };
            yield return new object[] { kind, IWorkDurationUnit.Second, "[s].000\"s\" ", "-129661.750s " };
        }
    }

    [Theory]
    [MemberData(nameof(MillisecondDurationContainers))]
    public void Fixed_millisecond_ranges_preserve_raw_seconds_and_saved_destination_display(
        IWorkDocumentKind kind, IWorkDurationUnit largest, string code, string expected) {
        using var package = DurationFormatPackage(kind, -129661.75, DurationFormat(largest: (ulong)largest, smallest: 32));
        var source = IWorkSourceDocument.Open(package, kind);
        var cell = Assert.Single(ReadSelectedRichTable(source, kind).Item1.Cells);
        Assert.Equal(-129661.75d, cell.Value);
        Assert.Equal(IWorkCellKind.Duration, cell.ValueKind);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        Assert.Equal(largest, cell.NumberFormat!.DurationFormat!.LargestUnit);
        Assert.Equal(IWorkDurationUnit.Millisecond, cell.NumberFormat.DurationFormat.SmallestUnit);
        Assert.Equal(code, cell.NumberFormat.ToSpreadsheetFormatCode());
        Assert.True(cell.TryGetFormattedNumber(out string display, out _));
        Assert.Equal(expected, display);
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult();
            result.Report.RequireCompleteEditableReconstruction();
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_DURATION_DISPLAY_APPROXIMATED"
                && d.LossKind == OfficeConversionLossKind.Approximation);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            Assert.Equal(-129661.75d / 86400, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
            Assert.Equal(code, reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
            Assert.Equal(expected, Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
        } else if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult(new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_NUMBER_FORMAT_APPROXIMATED"
                && d.LossKind == OfficeConversionLossKind.Approximation);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Word.WordDocument.Load(saved);
            Assert.Equal(expected, reopened.Tables[0].Rows[0].Cells[0].Paragraphs[0].Text);
            Assert.Empty(reopened.ValidateDocument());
        } else {
            using var result = source.ToPowerPointPresentationResult();
            result.Report.RequireCompleteEditableReconstruction();
            Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_KEYNOTE_NUMBER_FORMAT_APPROXIMATED"
                && d.LossKind == OfficeConversionLossKind.Approximation);
            result.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.PowerPoint.PowerPointPresentation.Load(saved);
            Assert.Equal(expected, Assert.Single(reopened.Slides[0].Tables).GetCell(0, 0).Text);
            Assert.Empty(reopened.ValidateDocument());
        }
    }

    [Theory]
    [InlineData(IWorkDurationUnit.Hour, "[h]\"h\" m\"m\" s.000\"s\" ")]
    [InlineData(IWorkDurationUnit.Minute, "[m]\"m\" s.000\"s\" ")]
    [InlineData(IWorkDurationUnit.Second, "[s].000\"s\" ")]
    public void Blank_millisecond_cells_preserve_selected_native_Xlsx_precision(IWorkDurationUnit largest, string code) {
        using var package = EmptyDurationFormatPackage(IWorkDocumentKind.Numbers,
            format: DurationFormat(largest: (ulong)largest, smallest: 32));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(code, reopened.Sheets[0].CellAt(1, 1).GetStyle().NumberFormatCode);
        Assert.Equal("", Assert.Single(reopened.Sheets[0].Range("A1").CreateVisualSnapshot().Cells).Text);
    }

    [Theory]
    [InlineData(IWorkDurationUnit.Hour, 59.9995, "0h 1m 0.000s ")]
    [InlineData(IWorkDurationUnit.Hour, 3599.9995, "1h 0m 0.000s ")]
    [InlineData(IWorkDurationUnit.Minute, -59.9995, "-1m 0.000s ")]
    [InlineData(IWorkDurationUnit.Second, -0.9995, "-1.000s ")]
    public void Millisecond_precision_uses_shared_carry_before_component_display(
        IWorkDurationUnit largest, double seconds, string expected) {
        using var package = DurationFormatPackage(IWorkDocumentKind.Numbers, seconds,
            DurationFormat(largest: (ulong)largest, smallest: 32));
        var cell = Assert.Single(IWorkSourceDocument.Open(package).ReadNumbers().Sheets[0].Tables[0].Cells);
        Assert.True(cell.TryGetFormattedNumber(out string display, out _));
        Assert.Equal(expected, display);
        package.Position = 0;
        var read = new OfficeDocumentReaderBuilder().AddIWorkHandler().Build().ReadDocument(package, "milliseconds.numbers");
        Assert.Equal(expected, Assert.Single(Assert.Single(read.Tables).Rows)[0]);
        Assert.Contains(read.Diagnostics, d => d.Code == "IWORK_READER_NUMBER_FORMAT_APPROXIMATED");
    }
}
