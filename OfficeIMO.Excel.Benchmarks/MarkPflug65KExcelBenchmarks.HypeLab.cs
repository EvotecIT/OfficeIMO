using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using System.Globalization;

namespace OfficeIMO.Excel.Benchmarks;

public partial class MarkPflug65KXlsxBenchmarks {
    /// <summary>Reads a fresh sheet and consumes all fourteen typed sales columns.</summary>
    [Benchmark]
    public ExcelReadObservation HypeLab() {
        var sheet = global::HypeLab.IO.Excel.ExcelReader.ExtractSheetData(MarkPflug65KFixture.XlsxPath);
        if (!sheet.Headers.SequenceEqual(Headers))
            throw new InvalidDataException("The sales workbook headers do not match the typed scan.");
        var observation = new ExcelObservationAccumulator();
        foreach (string?[] row in sheet.Rows) {
            if (row.Length < Headers.Length) throw new InvalidDataException("A sales row is incomplete.");
            observation.BeginRow();
            for (int column = 0; column < 5; column++)
                observation.Add(row[column] ?? throw new InvalidDataException("A sales text value is missing."));
            observation.Add(ReadHypeLabDate(row[5]));
            observation.Add(int.Parse(row[6]!, CultureInfo.InvariantCulture));
            observation.Add(ReadHypeLabDate(row[7]));
            observation.Add(int.Parse(row[8]!, CultureInfo.InvariantCulture));
            // String-based readers still observe the same IEEE-to-decimal values.
            for (int column = 9; column < Headers.Length; column++)
                observation.Add((decimal)double.Parse(row[column]!, NumberStyles.Float, CultureInfo.InvariantCulture));
        }
        return Validate(nameof(HypeLab), observation.Build());
    }

    private static DateTime ReadHypeLabDate(string? value) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double serial)
        ? DateTime.FromOADate(serial)
        : DateTime.Parse(value!, CultureInfo.InvariantCulture);
}
