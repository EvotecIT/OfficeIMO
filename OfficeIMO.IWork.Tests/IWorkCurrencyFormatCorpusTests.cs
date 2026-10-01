using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkCurrencyFormatCorpusTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "numbers-parser", "currency-formats." + extension);

    [Fact]
    public void Independent_currency_formats_retain_numeric_values_and_source_semantics() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Fixture("numbers")))).ToLowerInvariant());
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(Fixture("numbers")).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(manifest.RootElement.GetProperty("cases").GetArrayLength(), table.RowCount);
        Assert.DoesNotContain(projection.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            Assert.Equal(expected.GetProperty("label").GetString(), table.GetCell(row, 1)!.Value);
            IWorkTableCell cell = table.GetCell(row, 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("value").GetDouble(), Assert.IsType<double>(cell.Value), 12);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            Assert.Equal(IWorkNumberFormatKind.Currency, format.Kind);
            Assert.Equal(expected.GetProperty("currencyCode").GetString(), format.CurrencyCode);
            JsonElement decimals = expected.GetProperty("decimalPlaces");
            Assert.Equal(decimals.ValueKind == JsonValueKind.Null ? (int?)null : decimals.GetInt32(), format.DecimalPlaces);
            Assert.Equal(expected.GetProperty("thousandsSeparator").GetBoolean(), format.ThousandsSeparator);
            Assert.Equal((IWorkNegativeNumberStyle)expected.GetProperty("negativeStyle").GetInt32(), format.NegativeStyle);
            Assert.Equal(expected.GetProperty("useAccountingStyle").GetBoolean(), format.UseAccountingStyle);
        }
    }

    [Fact]
    public void Saved_currency_amounts_keep_identifiers_and_explicit_display_approximation() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"));
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_CURRENCY_DISPLAY_APPROXIMATED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:B11").CreateVisualSnapshot();
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            ExcelCellData cell = sheet.CellAt(row, 2).GetValue();
            Assert.Equal(ExcelCellDataKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("value").GetDouble(), Assert.IsType<double>(cell.Value), 12);
            var visual = Assert.Single(snapshot.Cells, c => c.Row == row);
            Assert.Equal(expected.GetProperty("destinationDisplayText").GetString(), visual.Text);
            if (expected.GetProperty("negativeStyle").GetInt32() is 1 or 3) Assert.Equal("FF0000", visual.Style.FontColorHex);
        }
    }
}
