using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using CsvHelper.Configuration;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using CsvHelperReader = CsvHelper.CsvReader;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Owned-string traversal of the 50,000-line NCsvPerf PackageAssets input in its plain and quoted
/// shapes. Every reader decodes and unescapes all 25 fields from the same in-memory UTF-8 bytes.
/// </summary>
[MemoryDiagnoser]
public class PackageAssetsCsvBenchmarks {
    private byte[] _input = Array.Empty<byte>();
    private CsvReadObservation _expected;

    [Params(PackageAssetsCsvShape.Plain, PackageAssetsCsvShape.Quotes)]
    public PackageAssetsCsvShape Shape { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _input = PackageAssetsCsvFixture.CreateInput(Shape);
        _expected = PackageAssetsCsvObservation.Expected(Shape);
        OfficeIMO(); Sep(); Sylvan(); CsvHelper();
    }

    [Benchmark(Baseline = true)]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenReader(new MemoryStream(_input, writable: false),
            new CsvReaderOptions { DetectDelimiter = false, HasHeaderRow = false });
        return Validate(PackageAssetsCsvObservation.Observe(reader));
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader(options => options with { HasHeader = false, Unescape = true })
            .From(new MemoryStream(_input, writable: false));
        var values = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            int count = row.ColCount;
            for (int column = 0; column < count; column++) values.Add(row[column].ToString());
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation Sylvan() {
        using var text = new StreamReader(new MemoryStream(_input, writable: false));
        using SylvanCsvDataReader reader = SylvanCsvDataReader.Create(text, new Sylvan.Data.Csv.CsvDataReaderOptions { HasHeaders = false });
        return Validate(PackageAssetsCsvObservation.Observe(reader));
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        using var text = new StreamReader(new MemoryStream(_input, writable: false));
        using var reader = new CsvHelperReader(text, new CsvConfiguration(CultureInfo.InvariantCulture) { HasHeaderRecord = false });
        return Validate(PackageAssetsCsvObservation.Observe(reader));
    }

    private CsvReadObservation Validate(CsvReadObservation actual) => actual == _expected
        ? actual : throw new InvalidDataException($"Incorrect PackageAssets {Shape} traversal: {actual}; expected {_expected}.");
}

/// <summary>
/// Spaces-and-quotes PackageAssets input read with whitespace outside quotes trimmed and quotes
/// removed. Sylvan has no trimming option, so it is not part of this comparison.
/// </summary>
[MemoryDiagnoser]
public class PackageAssetsSpacesQuotesCsvBenchmarks {
    private byte[] _input = Array.Empty<byte>();
    private CsvReadObservation _expected;

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _input = PackageAssetsCsvFixture.CreateInput(PackageAssetsCsvShape.SpacesQuotes);
        _expected = PackageAssetsCsvObservation.Expected(PackageAssetsCsvShape.SpacesQuotes);
        OfficeIMO(); Sep(); CsvHelper();
    }

    [Benchmark(Baseline = true)]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenReader(new MemoryStream(_input, writable: false),
            new CsvReaderOptions { DetectDelimiter = false, HasHeaderRow = false, TrimWhitespace = true });
        return Validate(PackageAssetsCsvObservation.Observe(reader));
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader(options => options with { HasHeader = false, Trim = SepTrim.Outer, Unescape = true })
            .From(new MemoryStream(_input, writable: false));
        var values = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            int count = row.ColCount;
            for (int column = 0; column < count; column++) values.Add(row[column].ToString());
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation CsvHelper() {
        using var text = new StreamReader(new MemoryStream(_input, writable: false));
        using var reader = new CsvHelperReader(text, new CsvConfiguration(CultureInfo.InvariantCulture) {
            HasHeaderRecord = false,
            TrimOptions = TrimOptions.Trim
        });
        return Validate(PackageAssetsCsvObservation.Observe(reader));
    }

    private CsvReadObservation Validate(CsvReadObservation actual) => actual == _expected
        ? actual : throw new InvalidDataException($"Incorrect PackageAssets spaces-and-quotes traversal: {actual}; expected {_expected}.");
}

internal static class PackageAssetsCsvObservation {
    internal static CsvReadObservation Expected(PackageAssetsCsvShape shape) {
        var values = new CsvObservationAccumulator();
        int rows = 0;
        foreach (string[] row in PackageAssetsCsvFixture.ExpectedRows(shape)) {
            foreach (string value in row) values.Add(value);
            rows++;
        }
        return values.ToObservation(rows);
    }

    // Inlined into each benchmark so every reader gets its own monomorphic call sites. A shared body would let
    // tiered PGO devirtualize GetString only for whichever reader the setup profiled first.
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static CsvReadObservation Observe(DbDataReader reader) {
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.FieldCount;
            for (int column = 0; column < count; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return values.ToObservation(rows);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static CsvReadObservation Observe(CsvHelperReader reader) {
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.Parser.Count;
            for (int column = 0; column < count; column++) values.Add(reader.GetField(column));
            rows++;
        }
        return values.ToObservation(rows);
    }
}
