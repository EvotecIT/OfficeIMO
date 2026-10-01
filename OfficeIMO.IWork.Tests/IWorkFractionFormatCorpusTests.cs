using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkFractionFormatCorpusTests {
    private static string Fixture(string extension) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", "numbers-parser", "fraction-formats." + extension);

    [Fact]
    public void Independent_fraction_modes_retain_numeric_values_and_source_precision() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        Assert.Equal(manifest.RootElement.GetProperty("sourceSha256").GetString(),
            Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(Fixture("numbers")))).ToLowerInvariant());
        IWorkTable table = Assert.Single(Assert.Single(IWorkSourceDocument.Open(Fixture("numbers")).ReadNumbers().Sheets).Tables);
        Assert.Equal(manifest.RootElement.GetProperty("cases").GetArrayLength(), table.RowCount);
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            Assert.Equal(expected.GetProperty("label").GetString(), table.GetCell(row, 1)!.Value);
            IWorkTableCell cell = table.GetCell(row, 2)!;
            Assert.Equal(IWorkCellKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("portableValue").GetDouble(), Assert.IsType<double>(cell.Value));
            bool approximate = expected.GetProperty("numericValueIsApproximate").GetBoolean();
            Assert.Equal(approximate, cell.NumericValueIsApproximate);
            Assert.Equal(approximate ? expected.GetProperty("sourceNumberText").GetString() : null, cell.SourceNumberText);
            IWorkNumberFormat format = Assert.IsType<IWorkNumberFormat>(cell.NumberFormat);
            Assert.Equal(IWorkNumberFormatKind.Fraction, format.Kind);
            Assert.Equal(Enum.Parse<IWorkFractionAccuracy>(expected.GetProperty("fractionAccuracy").GetString()!), format.FractionAccuracy);
            Assert.Null(format.DecimalPlaces);
            Assert.False(format.ThousandsSeparator);
            Assert.Equal(IWorkNegativeNumberStyle.Minus, format.NegativeStyle);
            Assert.Null(format.CurrencyCode);
            Assert.False(format.UseAccountingStyle);
        }
    }

    [Fact]
    public void Saved_fraction_output_preserves_numeric_types_and_reports_display_approximation() {
        using JsonDocument manifest = JsonDocument.Parse(File.ReadAllText(Fixture("json")));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(Fixture("numbers"));
        Assert.False(result.IsVisualFallback);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED");
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_FRACTION_DISPLAY_APPROXIMATED"
            && d.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        ExcelRangeVisualSnapshot snapshot = sheet.Range($"B1:B{manifest.RootElement.GetProperty("cases").GetArrayLength()}").CreateVisualSnapshot();
        foreach (JsonElement expected in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            int row = expected.GetProperty("row").GetInt32();
            ExcelCellData cell = sheet.CellAt(row, 2).GetValue();
            Assert.Equal(ExcelCellDataKind.Number, cell.Kind);
            Assert.Equal(expected.GetProperty("portableValue").GetDouble(), Assert.IsType<double>(cell.Value));
            string text = Assert.Single(snapshot.Cells, c => c.Row == row).Text;
            Assert.Equal(expected.GetProperty("destinationDisplayText").GetString(), text);
            if (expected.GetProperty("sourceDisplayIsQualified").GetBoolean())
                Assert.Equal(expected.GetProperty("sourceDisplayText").GetString(), text);
            Assert.Equal(expected.GetProperty("destinationFormatCode").GetString(), sheet.CellAt(row, 2).GetStyle().NumberFormatCode);
        }
    }
}
