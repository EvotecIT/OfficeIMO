using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkDurationFormatCorpusTests {
    [Fact]
    public void Native_duration_format_preserves_source_cache_and_matches_paired_saved_xlsx() {
        string root = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(root, "numbers-parser", "duration-format.json")));
        var expected = manifest.RootElement;
        string sourcePath = Path.Combine(root, expected.GetProperty("sourceFixture").GetString()!);
        string nativePath = Path.Combine(root, expected.GetProperty("nativeExport").GetString()!);
        Assert.Equal(expected.GetProperty("sourceSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(sourcePath))).ToLowerInvariant());
        Assert.Equal(expected.GetProperty("nativeExportSha256").GetString(), Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(nativePath))).ToLowerInvariant());
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(sourcePath,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        var cell = result.Projection.Sheets[0].Tables[0].GetCell(5, 3)!;
        Assert.Equal(expected.GetProperty("source").GetProperty("seconds").GetDouble(), cell.Value);
        Assert.Equal(IWorkCellKind.Duration, cell.ValueKind);
        Assert.Equal(IWorkCellUnsupportedFeatures.None, cell.UnsupportedFeatures);
        var format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
        Assert.Equal(IWorkNumberFormatKind.Duration, format.Kind);
        Assert.Equal(IWorkDurationUnit.Hour, format.DurationFormat!.LargestUnit);
        Assert.Equal(IWorkDurationUnit.Minute, format.DurationFormat.SmallestUnit);
        Assert.Equal(IWorkDurationStyle.Abbreviated, format.DurationFormat.Style);
        Assert.False(format.DurationFormat.UseAutomaticUnits);
        Assert.True(cell.FormulaIsComplete);
        Assert.True(cell.TryGetFormattedNumber(out string sourceDisplay, out _));
        Assert.Equal(expected.GetProperty("source").GetProperty("producerDisplayText").GetString(), sourceDisplay);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        using var native = ExcelDocument.Load(nativePath);
        var actual = reopened.Sheets[0].CellAt(5, 3);
        var oracle = native.Sheets[0].CellAt(6, 3);
        Assert.Equal(expected.GetProperty("native").GetProperty("valueDays").GetDouble(), actual.GetValue<double>());
        Assert.Equal(oracle.GetValue<double>(), actual.GetValue<double>());
        Assert.Equal(oracle.GetStyle().NumberFormatCode, actual.GetStyle().NumberFormatCode);
        Assert.Equal(expected.GetProperty("native").GetProperty("formatCode").GetString(), actual.GetStyle().NumberFormatCode);
        Assert.Equal(Assert.Single(native.Sheets[0].Range("C6").CreateVisualSnapshot().Cells).Text,
            Assert.Single(reopened.Sheets[0].Range("C5").CreateVisualSnapshot().Cells).Text);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_CELL_FEATURES_UNASSESSED");
        Assert.True(result.Report.IsPartialEditableReconstruction);
    }
}
