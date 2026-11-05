using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using System.Data.Common;
using System.Globalization;
using SepLib = nietras.SeparatedValues.Sep;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Compares complete owned-string traversals through the current reader APIs.</summary>
[MemoryDiagnoser]
public class MarkPflug65KModernCsvBenchmarks {
    private CsvReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.CsvFileName);
        _expected = new CsvReadObservation(MarkPflug65KFixture.ExpectedRows,
            MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns,
            MarkPflug65KFixture.ExpectedCsvCharacters, MarkPflug65KFixture.ExpectedCsvChecksum);
        OfficeIMO(); OfficeIMO_InternStrings(); Sep(); CsvHelper(); FourLambda();
    }

    [Benchmark]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenReader(MarkPflug65KFixture.CsvPath,
            options: new CsvReaderOptions { DetectDelimiter = false });
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    /// <summary>Measures the existing per-read string reuse option with the same owned values.</summary>
    [Benchmark]
    public CsvReadObservation OfficeIMO_InternStrings() {
        using DbDataReader reader = CsvDocument.OpenReader(MarkPflug65KFixture.CsvPath,
            options: new CsvReaderOptions { DetectDelimiter = false, InternStrings = true });
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader().FromFile(MarkPflug65KFixture.CsvPath);
        var values = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            for (int column = 0; column < row.ColCount; column++) values.Add(row[column].ToString());
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        using var input = File.OpenRead(MarkPflug65KFixture.CsvPath);
        using var reader = global::CsvHelper.CsvReader.Create(input, new global::CsvHelper.CsvReaderOptions {
            FormatProvider = CultureInfo.InvariantCulture
        });
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation FourLambda() {
        using var input = File.OpenRead(MarkPflug65KFixture.CsvPath);
        using var reader = global::FourLambda.Csv.CsvReader.Create(input, hasHeaders: true);
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.ReadNext()) {
            for (int column = 0; column < reader.FieldCount; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    private CsvReadObservation Validate(CsvReadObservation actual) => actual == _expected
        ? actual : throw new InvalidDataException($"Incorrect decoded CSV traversal: {actual}.");
}
