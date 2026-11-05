using BenchmarkDotNet.Attributes;
using ExcelPRIME;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

public partial class MarkPflug65KXlsxBenchmarks {
    [Benchmark]
    public ExcelReadObservation ExcelPrime() => ReadExcelPrime(styledTypes: false);

    [Benchmark]
    public ExcelReadObservation ExcelPrime_TypedStyles() => ReadExcelPrime(styledTypes: true);

    private ExcelReadObservation ReadExcelPrime(bool styledTypes) {
        using var workbook = new Excel_PRIME();
        workbook.Open(MarkPflug65KFixture.XlsxPath, ExcelPrimeSalesObservation.CreateOptions(styledTypes));
        return Validate(nameof(ExcelPrime), ExcelPrimeSalesObservation.Observe(workbook));
    }
}

public partial class MarkPflug65KXlsbBenchmarks {
    [Benchmark]
    public ExcelReadObservation ExcelPrime() => ReadExcelPrime(styledTypes: false);

    [Benchmark]
    public ExcelReadObservation ExcelPrime_TypedStyles() => ReadExcelPrime(styledTypes: true);

    private ExcelReadObservation ReadExcelPrime(bool styledTypes) {
        using var workbook = new Excel_PRIMEXlsb();
        workbook.Open(MarkPflug65KFixture.XlsbPath, ExcelPrimeSalesObservation.CreateOptions(styledTypes));
        return Validate(nameof(ExcelPrime), ExcelPrimeSalesObservation.Observe(workbook));
    }
}

/// <summary>Consumes the same typed sales columns through either PRIME format reader.</summary>
internal static class ExcelPrimeSalesObservation {
    /// <summary>Runs the ordinary complete observations without collecting timings.</summary>
    internal static void ValidateReaders() {
        var xlsx = new MarkPflug65KXlsxBenchmarks();
        xlsx.Setup();
        _ = xlsx.ExcelPrime();
        _ = xlsx.ExcelPrime_TypedStyles();
        var xlsb = new MarkPflug65KXlsbBenchmarks();
        xlsb.Setup();
        _ = xlsb.ExcelPrime();
        _ = xlsb.ExcelPrime_TypedStyles();
        Console.WriteLine("PASS | Both PRIME readers and conversion options produce the complete expected typed observation.");
    }

    internal static Options CreateOptions(bool styledTypes) => new() {
        CellConversionType = styledTypes ? CellConversion.ExcelCellStyle : CellConversion.None
    };

    internal static ExcelReadObservation Observe(Excel_PRIME workbook) {
        using var sheet = workbook.GetSheet(workbook.SheetNames().First())
            ?? throw new InvalidDataException("The sales workbook has no first sheet.");
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        foreach (var row in sheet.GetRowData()) {
            using (row) {
                if (header) { header = false; continue; }
                if (row is null) throw new InvalidDataException("The sales workbook contains a missing row.");
                observation.BeginRow();
                for (int column = 1; column <= 5; column++)
                    observation.Add(row.GetCell(column).CellValue.ToString()
                        ?? throw new InvalidDataException("The sales workbook contains a missing text cell."));
                observation.Add(row.GetCell(6).CellValue.AsDateTime);
                observation.Add(row.GetCell(7).CellValue.AsInt32);
                observation.Add(row.GetCell(8).CellValue.AsDateTime);
                observation.Add(row.GetCell(9).CellValue.AsInt32);
                // All numeric-cell lanes observe the same IEEE-to-decimal values.
                for (int column = 10; column <= 14; column++)
                    observation.Add((decimal)row.GetCell(column).CellValue.AsDouble);
            }
        }
        return observation.Build();
    }
}
