using System.Globalization;
using System.Data.Common;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using ExcelReader.Core.Reader;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using CsvHelperReader = CsvHelper.CsvReader;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using LumenWorksCsvReader = CsvReader.CsvReader;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Neutral all-field scan of the hash-pinned 65K sales CSV. Every compatible library
/// performs the same decoded-string traversal and must produce the same observation.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class MarkPflug65KCsvBenchmarks {
    private CsvReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.CsvFileName);
        _expected = new CsvReadObservation(
            MarkPflug65KFixture.ExpectedRows,
            MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns,
            MarkPflug65KFixture.ExpectedCsvCharacters,
            MarkPflug65KFixture.ExpectedCsvChecksum);
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(ExcelReaderNet), ExcelReaderNet());
        Validate(nameof(Sep), Sep());
        Validate(nameof(Sylvan), Sylvan());
        Validate(nameof(CsvHelper), CsvHelper());
        Validate(nameof(DataplatDbatools), DataplatDbatools());
        Validate(nameof(LumenWorks), LumenWorks());
    }

    [Benchmark]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenReader(MarkPflug65KFixture.CsvPath,
            options: new CsvReaderOptions { DetectDelimiter = false });
        return Observe(reader);
    }

    [Benchmark]
    public CsvReadObservation ExcelReaderNet() {
        using ExcelReaderNetCsvReader reader = ExcelReaderApi.FromCsvFile(MarkPflug65KFixture.CsvPath);
        var values = new CsvObservationAccumulator();
        int rows = 0;
        bool header = true;
#if NET10_0_OR_GREATER
        foreach (Row row in reader.FirstSheet) {
#else
        foreach (Row row in reader) {
#endif
            if (header) {
                header = false;
                continue;
            }

            int count = row.ColumnCount;
            for (int column = 0; column < count; column++) values.Add(row[column].GetString());
            rows++;
        }

        return values.ToObservation(rows);
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader().FromFile(MarkPflug65KFixture.CsvPath);
        var values = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            int count = row.ColCount;
            for (int column = 0; column < count; column++) values.Add(row[column].ToString());
            rows++;
        }

        return values.ToObservation(rows);
    }

    [Benchmark]
    public CsvReadObservation Sylvan() {
        using var text = new StreamReader(MarkPflug65KFixture.CsvPath);
        using SylvanCsvDataReader reader = SylvanCsvDataReader.Create(text);
        return Observe(reader);
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true };
        using var text = new StreamReader(MarkPflug65KFixture.CsvPath);
        using var reader = new CsvHelperReader(text, configuration);
        reader.Read();
        reader.ReadHeader();

        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.Parser.Count;
            for (int column = 0; column < count; column++) values.Add(reader.GetField(column));
            rows++;
        }

        return values.ToObservation(rows);
    }

    [Benchmark]
    public CsvReadObservation DataplatDbatools() {
        using var reader = new DataplatCsvDataReader(MarkPflug65KFixture.CsvPath);
        return Observe(
            reader.FieldCount,
            read: reader.Read,
            value: ordinal => Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture) ?? string.Empty);
    }

    [Benchmark]
    public CsvReadObservation LumenWorks() {
        using var text = new StreamReader(MarkPflug65KFixture.CsvPath);
        using var reader = new LumenWorksCsvReader(text, hasHeaders: true);
        return Observe(reader.FieldCount, read: reader.ReadNextRecord, value: ordinal => reader[ordinal]);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static CsvReadObservation Observe(
        int fieldCount,
        Func<bool> read,
        Func<int, string> value) {
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (read()) {
            for (int column = 0; column < fieldCount; column++) values.Add(value(column));
            rows++;
        }

        return values.ToObservation(rows);
    }

    // Inlined into each benchmark so every reader gets its own monomorphic call sites. A shared body would let
    // tiered PGO devirtualize GetString only for whichever reader the setup profiled first.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static CsvReadObservation Observe(DbDataReader reader) {
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.FieldCount;
            for (int column = 0; column < count; column++) values.Add(reader.GetString(column));
            rows++;
        }

        return values.ToObservation(rows);
    }

    private void Validate(string library, CsvReadObservation actual) {
        if (actual != _expected
            || actual.Rows != MarkPflug65KFixture.ExpectedRows
            || actual.Cells != MarkPflug65KFixture.ExpectedRows * MarkPflug65KFixture.ExpectedColumns) {
            throw new InvalidDataException(
                $"{library} did not perform the same CSV workload. Expected {_expected}; actual {actual}.");
        }
    }
}
