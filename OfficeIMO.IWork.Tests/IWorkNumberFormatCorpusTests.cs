using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNumberFormatCorpusTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "numbers-parser", "number-formats." + extension);

    [Fact]
    public void Independent_number_formats_retain_the_producer_values_and_metadata() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        string hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Fixture("numbers")))).ToLowerInvariant();
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(), hash);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(Fixture("numbers")).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        JsonElement cases = manifest.RootElement.GetProperty("cases");
        Assert.Equal(cases.GetArrayLength(), table.RowCount);
        Assert.Equal(2, table.ColumnCount);
        foreach (JsonElement expected in cases.EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            Assert.Equal(expected.GetProperty("label").GetString(), table.GetCell(row, 1)!.Value);
            IWorkTableCell cell = table.GetCell(row, 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("value").GetDouble(), Assert.IsType<double>(cell.Value), 12);
            bool approximate = expected.GetProperty("numericValueIsApproximate").GetBoolean();
            Assert.Equal(approximate ? expected.GetProperty("sourceNumberText").GetString() : null, cell.SourceNumberText);
            Assert.Equal(approximate, cell.NumericValueIsApproximate);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            Assert.Equal(Enum.Parse<IWorkNumberFormatKind>(expected.GetProperty("kind").GetString()!), format.Kind);
            JsonElement decimals = expected.GetProperty("decimalPlaces");
            Assert.Equal(decimals.ValueKind == JsonValueKind.Null ? (int?)null : decimals.GetInt32(), format.DecimalPlaces);
            Assert.Equal(expected.GetProperty("thousandsSeparator").GetBoolean(), format.ThousandsSeparator);
            Assert.Equal((IWorkNegativeNumberStyle)expected.GetProperty("negativeStyle").GetInt32(), format.NegativeStyle);
        }
    }

    [Fact]
    public void Independent_number_formats_survive_saved_xlsx_and_match_producer_display_text() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"), conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMERIC_VALUE_APPROXIMATED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        int count = manifest.RootElement.GetProperty("cases").GetArrayLength();
        ExcelRangeVisualSnapshot snapshot = sheet.Range($"B1:B{count}").CreateVisualSnapshot();
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            ExcelCellData cell = sheet.CellAt(row, 2).GetValue();
            Assert.Equal(ExcelCellDataKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("value").GetDouble(), Assert.IsType<double>(cell.Value), 12);
            Assert.Equal(expected.GetProperty("displayText").GetString(), Assert.Single(snapshot.Cells, c => c.Row == row).Text);
            int negativeStyle = expected.GetProperty("negativeStyle").GetInt32();
            if (negativeStyle is 1 or 3) Assert.Contains("[Red]", sheet.CellAt(row, 2).GetStyle().NumberFormatCode!);
        }
    }

    [Fact]
    public void Independent_red_number_styles_reach_the_image_snapshot() {
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"), conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelRangeVisualSnapshot snapshot = reopened.Sheets[0].Range("B1:B13").CreateVisualSnapshot();
        foreach (int row in new[] { 3, 5, 8, 10 }) {
            Assert.Equal("FF0000", Assert.Single(snapshot.Cells, c => c.Row == row).Style.FontColorHex);
        }
    }
}
