using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkScientificFormatCorpusTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "numbers-parser", "scientific-formats." + extension);

    [Fact]
    public void Independent_scientific_formats_retain_values_precision_and_oracle_boundaries() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Fixture("numbers")))).ToLowerInvariant());
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(Fixture("numbers")).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(manifest.RootElement.GetProperty("cases").GetArrayLength(), table.RowCount);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            Assert.Equal(expected.GetProperty("label").GetString(), table.GetCell(row, 1)!.Value);
            IWorkTableCell cell = table.GetCell(row, 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            double value = Assert.IsType<double>(cell.Value);
            Assert.Equal(expected.GetProperty("portableValue").GetDouble(), value);
            bool approximate = expected.GetProperty("numericValueIsApproximate").GetBoolean();
            Assert.Equal(approximate, cell.NumericValueIsApproximate);
            Assert.Equal(approximate ? expected.GetProperty("sourceNumberText").GetString() : null, cell.SourceNumberText);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            Assert.Equal(IWorkNumberFormatKind.Scientific, format.Kind);
            JsonElement decimals = expected.GetProperty("decimalPlaces");
            Assert.Equal(decimals.ValueKind == JsonValueKind.Null ? (int?)null : decimals.GetInt32(), format.DecimalPlaces);
            Assert.Equal(IWorkNegativeNumberStyle.Minus, format.NegativeStyle);
            Assert.False(format.ThousandsSeparator);
            Assert.Null(format.CurrencyCode);
            Assert.False(format.UseAccountingStyle);
            bool qualified = expected.GetProperty("sourceDisplayIsQualified").GetBoolean();
            Assert.Equal(decimals.ValueKind != JsonValueKind.Null, qualified);
            if (qualified) Assert.Equal(expected.GetProperty("sourceDisplayText").GetString(), expected.GetProperty("destinationDisplayText").GetString());
        }
    }

    [Fact]
    public void Saved_scientific_values_match_the_explicit_producer_and_report_automatic_approximation() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"));
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED"
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
            Assert.Equal(expected.GetProperty("portableValue").GetDouble(), Assert.IsType<double>(cell.Value));
            Assert.Equal(expected.GetProperty("destinationDisplayText").GetString(), Assert.Single(snapshot.Cells, c => c.Row == row).Text);
            Assert.EndsWith("E+00", sheet.CellAt(row, 2).GetStyle().NumberFormatCode!);
        }
    }
}
