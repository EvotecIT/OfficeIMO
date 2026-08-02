using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using Sylvan.Data.Excel;
using SylvanExcelDataReader = Sylvan.Data.Excel.ExcelDataReader;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Neutral typed scan of the hash-pinned 65K sales BIFF8 XLS fixture. Every
/// compatible reader consumes the same fourteen columns and validated payload.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class MarkPflug65KXlsBenchmarks {
    private ExcelReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.XlsFileName);
        _expected = OfficeIMO();
        Validate(nameof(Sylvan), Sylvan());
        Validate(nameof(ExcelDataReader), ExcelDataReader());
    }

    [Benchmark]
    public ExcelReadObservation OfficeIMO() {
        using DbDataReader reader = ExcelDocument.OpenDataReader(
            MarkPflug65KFixture.XlsPath,
            new ExcelReadOptions { NumericAsDecimal = true });
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation Sylvan() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsPath);
        using SylvanExcelDataReader reader = SylvanExcelDataReader.Create(
            stream,
            ExcelWorkbookType.Excel,
            new ExcelDataReaderOptions { Schema = ExcelSchema.Default });
        return MarkPflug65KXlsxBenchmarks.Observe(reader);
    }

    [Benchmark]
    public ExcelReadObservation ExcelDataReader() {
        using var stream = File.OpenRead(MarkPflug65KFixture.XlsPath);
        using global::ExcelDataReader.IExcelDataReader reader =
            global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) {
            return default;
        }

        var observation = new ExcelObservationAccumulator();
        while (reader.Read()) {
            MarkPflug65KXlsxBenchmarks.AddRow(
                ref observation,
                ordinal => Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty,
                ordinal => MarkPflug65KXlsxBenchmarks.ReadDate(reader.GetValue(ordinal)),
                ordinal => Convert.ToInt32(reader.GetValue(ordinal), CultureInfo.InvariantCulture),
                ordinal => Convert.ToDecimal(reader.GetValue(ordinal), CultureInfo.InvariantCulture));
        }

        return observation.Build();
    }

    private void Validate(string library, ExcelReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same XLS workload. Expected {_expected}; actual {actual}.");
        }
    }
}
