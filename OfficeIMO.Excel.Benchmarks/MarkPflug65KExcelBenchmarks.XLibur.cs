using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

public partial class MarkPflug65KXlsxBenchmarks {
    /// <summary>Observes the same typed sales values through XLibur's loaded workbook API.</summary>
    [Benchmark]
    public ExcelReadObservation XLibur() {
        using var workbook = new global::XLibur.Excel.XLWorkbook(MarkPflug65KFixture.XlsxPath);
        var sheet = workbook.Worksheets.First();
        int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var observation = new ExcelObservationAccumulator();
        for (int row = 2; row <= lastRow; row++) {
            observation.BeginRow();
            for (int column = 1; column <= 5; column++) observation.Add(sheet.Cell(row, column).GetString());
            observation.Add(sheet.Cell(row, 6).GetDateTime());
            observation.Add(checked((int)sheet.Cell(row, 7).GetDouble()));
            observation.Add(sheet.Cell(row, 8).GetDateTime());
            observation.Add(checked((int)sheet.Cell(row, 9).GetDouble()));
            for (int column = 10; column <= 14; column++) observation.Add((decimal)sheet.Cell(row, column).GetDouble());
        }
        return Validate(nameof(XLibur), observation.Build());
    }
}
