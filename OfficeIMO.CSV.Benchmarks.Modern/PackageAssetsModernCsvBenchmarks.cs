using System.Data.Common;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using OfficeIMO.Benchmarks;
using SepLib = nietras.SeparatedValues.Sep;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Owned-string traversal of the 50,000-line NCsvPerf PackageAssets shapes through the current
/// reader APIs. Spaces-and-quotes input trims whitespace outside quotes and removes the quotes.
/// </summary>
[MemoryDiagnoser]
public class PackageAssetsModernCsvBenchmarks {
    private byte[] _input = Array.Empty<byte>();
    private CsvReadObservation _expected;

    [Params(PackageAssetsCsvShape.Plain, PackageAssetsCsvShape.Quotes, PackageAssetsCsvShape.SpacesQuotes)]
    public PackageAssetsCsvShape Shape { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        _input = PackageAssetsCsvFixture.CreateInput(Shape);
        var expected = new CsvObservationAccumulator();
        int rows = 0;
        foreach (string[] row in PackageAssetsCsvFixture.ExpectedRows(Shape)) {
            foreach (string value in row) expected.Add(value);
            rows++;
        }
        _expected = expected.ToObservation(rows);
        OfficeIMO(); Sep(); CsvHelper();
    }

    private bool Trim => Shape == PackageAssetsCsvShape.SpacesQuotes;

    [Benchmark(Baseline = true)]
    public CsvReadObservation OfficeIMO() {
        using DbDataReader reader = CsvDocument.OpenReader(new MemoryStream(_input, writable: false),
            new CsvReaderOptions { DetectDelimiter = false, HasHeaderRow = false, TrimWhitespace = Trim });
        var values = new CsvObservationAccumulator();
        int rows = 0;
        while (reader.Read()) {
            int count = reader.FieldCount;
            for (int column = 0; column < count; column++) values.Add(reader.GetString(column));
            rows++;
        }
        return Validate(values.ToObservation(rows));
    }

    [Benchmark]
    public CsvReadObservation Sep() {
        using var reader = SepLib.Reader(options => options with {
            HasHeader = false,
            Trim = Trim ? SepTrim.Outer : SepTrim.None,
            Unescape = true
        }).From(new MemoryStream(_input, writable: false));
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
        using var input = new MemoryStream(_input, writable: false);
        using var reader = global::CsvHelper.CsvReader.Create(input, new global::CsvHelper.CsvReaderOptions {
            FormatProvider = CultureInfo.InvariantCulture,
            HasHeader = false,
            TrimWhitespace = Trim
        });
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
        ? actual : throw new InvalidDataException($"Incorrect PackageAssets {Shape} traversal: {actual}; expected {_expected}.");
}
