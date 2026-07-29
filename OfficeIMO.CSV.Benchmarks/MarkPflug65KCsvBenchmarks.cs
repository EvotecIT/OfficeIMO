using System.Globalization;
using System.Data.Common;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using CsvHelperReader = CsvHelper.CsvReader;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using LumenWorksCsvReader = CsvReader.CsvReader;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Neutral all-field scan of the hash-pinned 65K sales CSV. Every compatible library
/// performs the same decoded-string traversal and must produce the same observation.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class MarkPflug65KCsvBenchmarks {
    private const ulong ChecksumOffset = 14695981039346656037UL;
    private const ulong ChecksumPrime = 1099511628211UL;
    private CsvReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        MarkPflug65KFixture.EnsureAuthentic();
        _expected = OfficeIMO();
        Validate(nameof(Sep), Sep());
        Validate(nameof(Sylvan), Sylvan());
        Validate(nameof(CsvHelper), CsvHelper());
        Validate(nameof(DataplatDbatools), DataplatDbatools());
        Validate(nameof(LumenWorks), LumenWorks());
    }

    [Benchmark]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenDataReader(
            MarkPflug65KFixture.CsvPath,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream, DetectDelimiter = false });
        return Observe(reader.FieldCount, read: reader.Read, value: reader.GetString);
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader().FromFile(MarkPflug65KFixture.CsvPath);
        int rows = 0;
        int cells = 0;
        long characters = 0;
        ulong checksum = ChecksumOffset;
        foreach (var row in reader) {
            rows++;
            for (int column = 0; column < row.ColCount; column++) {
                cells++;
                string decoded = row[column].ToString();
                characters += decoded.Length;
                AddValue(ref checksum, decoded);
            }
        }

        return new CsvReadObservation(rows, cells, characters, checksum);
    }

    [Benchmark]
    public CsvReadObservation Sylvan() {
        using var text = new StreamReader(MarkPflug65KFixture.CsvPath);
        using SylvanCsvDataReader reader = SylvanCsvDataReader.Create(text);
        return Observe(reader.FieldCount, read: reader.Read, value: reader.GetString);
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = true };
        using var text = new StreamReader(MarkPflug65KFixture.CsvPath);
        using var reader = new CsvHelperReader(text, configuration);
        reader.Read();
        reader.ReadHeader();

        int rows = 0;
        int cells = 0;
        long characters = 0;
        ulong checksum = ChecksumOffset;
        while (reader.Read()) {
            rows++;
            int count = reader.Parser.Count;
            for (int column = 0; column < count; column++) {
                cells++;
                string decoded = reader.GetField(column) ?? string.Empty;
                characters += decoded.Length;
                AddValue(ref checksum, decoded);
            }
        }

        return new CsvReadObservation(rows, cells, characters, checksum);
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

    private static CsvReadObservation Observe(
        int fieldCount,
        Func<bool> read,
        Func<int, string> value) {
        int rows = 0;
        int cells = 0;
        long characters = 0;
        ulong checksum = ChecksumOffset;
        while (read()) {
            rows++;
            for (int column = 0; column < fieldCount; column++) {
                cells++;
                string decoded = value(column);
                characters += decoded.Length;
                AddValue(ref checksum, decoded);
            }
        }

        return new CsvReadObservation(rows, cells, characters, checksum);
    }

    private static void AddValue(ref ulong checksum, string value) {
        foreach (char character in value) {
            checksum ^= character;
            checksum *= ChecksumPrime;
        }

        checksum ^= (ulong)value.Length;
        checksum *= ChecksumPrime;
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

public readonly record struct CsvReadObservation(
    int Rows,
    int Cells,
    long Characters,
    ulong Checksum);
