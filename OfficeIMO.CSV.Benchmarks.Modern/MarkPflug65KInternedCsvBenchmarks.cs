using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using System.Data.Common;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Measures the per-read string reuse option on the 65K sales CSV. It runs in its own lane so that the
/// library comparison's process never executes this mode before the default reader is measured.
/// </summary>
[MemoryDiagnoser]
public class MarkPflug65KInternedCsvBenchmarks {
    private CsvReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.CsvFileName);
        _expected = new CsvReadObservation(MarkPflug65KFixture.ExpectedRows,
            MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns,
            MarkPflug65KFixture.ExpectedCsvCharacters, MarkPflug65KFixture.ExpectedCsvChecksum);
        OfficeIMO_InternStrings();
    }

    [Benchmark]
    public CsvReadObservation OfficeIMO_InternStrings() {
        using DbDataReader reader = CsvDocument.OpenReader(MarkPflug65KFixture.CsvPath,
            options: new CsvReaderOptions { DetectDelimiter = false, InternStrings = true });
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.FieldCount;
            for (int column = 0; column < count; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    private CsvReadObservation Validate(CsvReadObservation actual) => actual == _expected
        ? actual : throw new InvalidDataException($"Incorrect decoded CSV traversal: {actual}.");
}
