using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using CsvHelper;
using CsvHelper.Configuration;
using nietras.SeparatedValues;
using OfficeIMO.Data;
using CsvHelperParser = CsvHelper.CsvParser;
using SepLib = nietras.SeparatedValues.Sep;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Compares equivalent outer-trim and quote-unescape behavior. The generated input uses
/// ASCII spaces because that is the common trimming contract supported by every lane.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvTrimUnescapeBenchmarks {
    private string _csvText = string.Empty;
    private CsvReadObservation _expected;

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var text = new StringBuilder(RowCount * 100);
        var expected = new CsvObservationAccumulator();
        for (int index = 1; index <= RowCount; index++) {
            string plain = string.Create(CultureInfo.InvariantCulture, $"alpha-{index}");
            string inside = string.Create(CultureInfo.InvariantCulture, $"  inside-{index}  ");
            string escaped = string.Create(CultureInfo.InvariantCulture, $"quote \"{index}\"");
            string region = string.Create(CultureInfo.InvariantCulture, $"region,{index}");

            text.Append("  ").Append(plain).Append("  ,  \"")
                .Append(inside.Replace("\"", "\"\"", StringComparison.Ordinal))
                .Append("\"  ,\"")
                .Append(escaped.Replace("\"", "\"\"", StringComparison.Ordinal))
                .Append("\",  , \"")
                .Append(region)
                .Append("\" \n");

            expected.Add(plain);
            expected.Add(inside);
            expected.Add(escaped);
            expected.Add(ReadOnlySpan<char>.Empty);
            expected.Add(region);
        }

        _csvText = text.ToString();
        _expected = expected.ToObservation(RowCount);
        Validate(nameof(OfficeIMODataReaderStrings), OfficeIMODataReaderStrings());
        Validate(nameof(SepStrings), SepStrings());
        Validate(nameof(CsvHelperStrings), CsvHelperStrings());
    }

    [Benchmark(Baseline = true)]
    public CsvReadObservation OfficeIMODataReaderStrings() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            _csvText,
            new CsvLoadOptions {
                DetectDelimiter = false,
                Delimiter = ',',
                HasHeaderRow = false,
                TrimWhitespace = true
            });
        var observation = new CsvObservationAccumulator();
        int rows = 0;
        int fieldCount = reader.FieldCount;
        while (reader.Read()) {
            rows++;
            for (int column = 0; column < fieldCount; column++) {
                observation.Add(reader.GetString(column));
            }
        }

        return observation.ToObservation(rows);
    }

    [Benchmark]
    public CsvReadObservation SepStrings() {
        var options = SepLib.New(',').Reader(value => value with {
            HasHeader = false,
            Unescape = true,
            Trim = SepTrim.Outer
        });
        using var reader = options.FromText(_csvText);
        var observation = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            rows++;
            for (int column = 0; column < row.ColCount; column++) {
                observation.Add(row[column].ToString());
            }
        }

        return observation.ToObservation(rows);
    }

    [Benchmark]
    public CsvReadObservation CsvHelperStrings() {
        using var text = new StringReader(_csvText);
        var configuration = new CsvConfiguration(CultureInfo.InvariantCulture) {
            HasHeaderRecord = false,
            TrimOptions = TrimOptions.Trim
        };
        using var parser = new CsvHelperParser(text, configuration);
        var observation = new CsvObservationAccumulator();
        int rows = 0;
        while (parser.Read()) {
            rows++;
            for (int column = 0; column < parser.Count; column++) {
                observation.Add(parser[column]);
            }
        }

        return observation.ToObservation(rows);
    }

    private void Validate(string library, CsvReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(
                $"{library} did not perform the same trim/unescape workload. Expected {_expected}; actual {actual}.");
        }
    }

}

/// <summary>
/// Compares equivalent transient-span trim and unescape contracts. Keeping this in a
/// separate benchmark type prevents zero-copy results from being ranked against string APIs.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvTrimUnescapeSpanBenchmarks {
    private string _csvText = string.Empty;
    private CsvReadObservation _expected;

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var text = new StringBuilder(RowCount * 100);
        var expected = new CsvObservationAccumulator();
        for (int index = 1; index <= RowCount; index++) {
            string plain = string.Create(CultureInfo.InvariantCulture, $"alpha-{index}");
            string inside = string.Create(CultureInfo.InvariantCulture, $"  inside-{index}  ");
            string escaped = string.Create(CultureInfo.InvariantCulture, $"quote \"{index}\"");
            string region = string.Create(CultureInfo.InvariantCulture, $"region,{index}");

            text.Append("  ").Append(plain).Append("  ,  \"")
                .Append(inside.Replace("\"", "\"\"", StringComparison.Ordinal))
                .Append("\"  ,\"")
                .Append(escaped.Replace("\"", "\"\"", StringComparison.Ordinal))
                .Append("\",  , \"")
                .Append(region)
                .Append("\" \n");

            expected.Add(plain);
            expected.Add(inside);
            expected.Add(escaped);
            expected.Add(ReadOnlySpan<char>.Empty);
            expected.Add(region);
        }

        _csvText = text.ToString();
        _expected = expected.ToObservation(RowCount);
        Validate(nameof(OfficeIMOFieldSpans), OfficeIMOFieldSpans());
        Validate(nameof(SepSpans), SepSpans());
    }

    [Benchmark(Baseline = true)]
    public CsvReadObservation OfficeIMOFieldSpans() {
        var visitor = new ObservingFieldSpanVisitor();
        CsvDocument.ReadFieldSpansFromText(
            _csvText,
            ref visitor,
            new CsvLoadOptions {
                DetectDelimiter = false,
                Delimiter = ',',
                HasHeaderRow = false,
                TrimWhitespace = true
            });
        return visitor.Observation;
    }

    [Benchmark]
    public CsvReadObservation SepSpans() {
        var options = SepLib.New(',').Reader(value => value with {
            HasHeader = false,
            Unescape = true,
            Trim = SepTrim.Outer
        });
        using var reader = options.FromText(_csvText);
        var observation = new CsvObservationAccumulator();
        int rows = 0;
        foreach (var row in reader) {
            rows++;
            for (int column = 0; column < row.ColCount; column++) {
                observation.Add(row[column].Span);
            }
        }

        return observation.ToObservation(rows);
    }

    private void Validate(string library, CsvReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(
                $"{library} did not perform the same span trim/unescape workload. Expected {_expected}; actual {actual}.");
        }
    }

    private struct ObservingFieldSpanVisitor : ICsvFieldSpanVisitor {
        private CsvObservationAccumulator _observation;
        private int _rows;

        internal CsvReadObservation Observation => _observation.ToObservation(_rows);

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value) {
            _rows = Math.Max(_rows, recordIndex + 1);
            _observation.Add(value);
        }
    }
}

/// <summary>
/// Compares equivalent explicit typed materialization. Both lanes resolve the same headers,
/// decode quotes, create the same objects, preserve row order, and pass a property-by-property
/// preflight before measurement.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvTypedSequentialBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(25_000, 100_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() => _fixture = CsvTypedMaterializationFixture.Create(RowCount);

    [Benchmark(Baseline = true)]
    public CsvBenchmarkRow[] OfficeIMOManual() => _fixture.OfficeIMOManual();

    [Benchmark]
    public CsvBenchmarkRow[] SepSequential() => _fixture.SepSequential();
}

/// <summary>
/// Measures the convenience cost of OfficeIMO's automatic property mapper against its
/// equivalent explicit typed-reader loop. This is intentionally not a competitor ranking.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvAutomaticMappingBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(25_000, 100_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() => _fixture = CsvTypedMaterializationFixture.Create(RowCount);

    [Benchmark(Baseline = true)]
    public CsvBenchmarkRow[] OfficeIMOManual() => _fixture.OfficeIMOManual();

    [Benchmark]
    public CsvBenchmarkRow[] OfficeIMORowsAs() => _fixture.OfficeIMORowsAs();
}

/// <summary>
/// Compares equivalent ordered parallel typed materialization. Both lanes parse the same
/// source, create the same objects, preserve row order, and use the same worker limit.
/// Both lanes use their public transient-record APIs so generic IDataRecord dispatch is not
/// misreported as parser cost.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvParallelScalingBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(16)]
    public int DegreeOfParallelism { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = CsvTypedMaterializationFixture.Create(100_000);
        _fixture.ValidateParallel(OfficeIMOParallel(), SepParallel());
    }

    [Benchmark(Baseline = true)]
    public CsvBenchmarkRow[] OfficeIMOParallel() => _fixture.OfficeIMORecordParallel(DegreeOfParallelism);

    [Benchmark]
    public CsvBenchmarkRow[] SepParallel() => _fixture.SepParallel(DegreeOfParallelism);
}

/// <summary>
/// Measures the smaller parallel crossover workload separately so fixed invocation counts can
/// keep each BDN iteration long enough without multiplying the 100,000-row workload.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvParallelCrossoverBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(16)]
    public int DegreeOfParallelism { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = CsvTypedMaterializationFixture.Create(25_000);
        _fixture.ValidateParallel(OfficeIMOParallel(), SepParallel());
    }

    [Benchmark(Baseline = true)]
    public CsvBenchmarkRow[] OfficeIMOParallel() => _fixture.OfficeIMORecordParallel(DegreeOfParallelism);

    [Benchmark]
    public CsvBenchmarkRow[] SepParallel() => _fixture.SepParallel(DegreeOfParallelism);
}

/// <summary>
/// Fixed-work production tuning for OfficeIMO's public transient-record path. This class has no
/// competitor ratio because batch size is an OfficeIMO implementation choice.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvParallelOfficeTuningBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(1024, 2048, 3072, 3584, 4096)]
    public int BatchSize { get; set; }

    [Params(16)]
    public int DegreeOfParallelism { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = CsvTypedMaterializationFixture.Create(100_000);
        _fixture.ValidateOfficeParallel(
            _fixture.OfficeIMORecordParallel(DegreeOfParallelism, BatchSize));
    }

    [Benchmark]
    public CsvBenchmarkRow[] OfficeIMORecordParallel() =>
        _fixture.OfficeIMORecordParallel(DegreeOfParallelism, BatchSize);
}

/// <summary>
/// Fixed-work production tuning for OfficeIMO's 25,000-row ordered-parallel crossover.
/// Kept separate from the sustained workload so ranks never compare different row counts.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class CsvParallelCrossoverTuningBenchmarks {
    private CsvTypedMaterializationFixture _fixture = null!;

    [Params(1024, 2048, 3072, 3584, 4096)]
    public int BatchSize { get; set; }

    [Params(16)]
    public int DegreeOfParallelism { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = CsvTypedMaterializationFixture.Create(25_000);
        _fixture.ValidateOfficeParallel(
            _fixture.OfficeIMORecordParallel(DegreeOfParallelism, BatchSize));
    }

    [Benchmark]
    public CsvBenchmarkRow[] OfficeIMORecordParallel() =>
        _fixture.OfficeIMORecordParallel(DegreeOfParallelism, BatchSize);
}

internal sealed class CsvTypedMaterializationFixture {
    private static readonly CultureInfo Invariant = CultureInfo.InvariantCulture;
    private CsvBenchmarkRow[] _expectedRows = [];
    private string _csvText = string.Empty;

    private int RowCount { get; set; }

    internal int TextLength => _csvText.Length;

    internal static CsvTypedMaterializationFixture Create(int rowCount) {
        var fixture = new CsvTypedMaterializationFixture { RowCount = rowCount };
        fixture.Setup();
        return fixture;
    }

    private void Setup() {
        _expectedRows = CsvBenchmarkData.Create(RowCount, CsvBenchmarkShape.Quoted);
        using var writer = new StringWriter(Invariant);
        CsvDocument.WriteObjects(
            writer,
            _expectedRows,
            new CsvSaveOptions { NewLine = "\n", DateTimeFormat = "O" });
        _csvText = writer.ToString();

        Validate(nameof(OfficeIMORowsAs), OfficeIMORowsAs());
        Validate(nameof(OfficeIMOManual), OfficeIMOManual());
        Validate(nameof(SepSequential), SepSequential());
        Validate(nameof(SepParallel), SepParallel(Environment.ProcessorCount));
    }

    internal CsvBenchmarkRow[] OfficeIMORowsAs() => ReadOfficeImoRows();

    internal CsvBenchmarkRow[] OfficeIMOManual() => ReadOfficeImoManualRows();

    internal CsvBenchmarkRow[] SepSequential() {
        using var reader = CreateSepReader();
        SepColumnMap columns = ResolveColumns(reader);
        return reader.Enumerate(row => ParseSepRow(row, columns)).ToArray();
    }

    internal CsvBenchmarkRow[] OfficeIMOParallel(int degreeOfParallelism, int batchSize = 2048) {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(_csvText);
        int id = reader.GetOrdinal(nameof(CsvBenchmarkRow.Id));
        int name = reader.GetOrdinal(nameof(CsvBenchmarkRow.Name));
        int department = reader.GetOrdinal(nameof(CsvBenchmarkRow.Department));
        int region = reader.GetOrdinal(nameof(CsvBenchmarkRow.Region));
        int isEnabled = reader.GetOrdinal(nameof(CsvBenchmarkRow.IsEnabled));
        int created = reader.GetOrdinal(nameof(CsvBenchmarkRow.Created));
        int score = reader.GetOrdinal(nameof(CsvBenchmarkRow.Score));
        int owner = reader.GetOrdinal(nameof(CsvBenchmarkRow.Owner));
        int ticketCount = reader.GetOrdinal(nameof(CsvBenchmarkRow.TicketCount));
        int notes = reader.GetOrdinal(nameof(CsvBenchmarkRow.Notes));
        return reader.RowsAsParallel(
            row => new CsvBenchmarkRow {
                Id = row.GetInt32(id),
                Name = row.GetString(name),
                Department = row.GetString(department),
                Region = row.GetString(region),
                IsEnabled = row.GetBoolean(isEnabled),
                Created = row.GetDateTime(created),
                Score = row.GetDecimal(score),
                Owner = row.GetString(owner),
                TicketCount = row.GetInt32(ticketCount),
                Notes = row.GetString(notes)
            },
            new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = batchSize
            }).ToArray();
    }

    internal CsvBenchmarkRow[] OfficeIMORecordParallel(int degreeOfParallelism, int? batchSize = null) =>
        CsvDocument.ReadTextRowsAsParallel<CsvBenchmarkRow>(
            _csvText,
            header => {
                int id = header.GetOrdinal(nameof(CsvBenchmarkRow.Id));
                int name = header.GetOrdinal(nameof(CsvBenchmarkRow.Name));
                int department = header.GetOrdinal(nameof(CsvBenchmarkRow.Department));
                int region = header.GetOrdinal(nameof(CsvBenchmarkRow.Region));
                int isEnabled = header.GetOrdinal(nameof(CsvBenchmarkRow.IsEnabled));
                int created = header.GetOrdinal(nameof(CsvBenchmarkRow.Created));
                int score = header.GetOrdinal(nameof(CsvBenchmarkRow.Score));
                int owner = header.GetOrdinal(nameof(CsvBenchmarkRow.Owner));
                int ticketCount = header.GetOrdinal(nameof(CsvBenchmarkRow.TicketCount));
                int notes = header.GetOrdinal(nameof(CsvBenchmarkRow.Notes));
                return row => new CsvBenchmarkRow {
                    Id = row.GetInt32(id),
                    Name = row.GetString(name),
                    Department = row.GetString(department),
                    Region = row.GetString(region),
                    IsEnabled = row.GetBoolean(isEnabled),
                    Created = row.GetDateTime(created),
                    Score = row.GetDecimal(score),
                    Owner = row.GetString(owner),
                    TicketCount = row.GetInt32(ticketCount),
                    Notes = row.GetString(notes)
                };
            },
            parallelOptions: new ParallelRowMappingOptions {
                MaxDegreeOfParallelism = degreeOfParallelism,
                BatchSize = batchSize
            }).ToArray();

    internal CsvBenchmarkRow[] SepParallel(int degreeOfParallelism) {
        using var reader = CreateSepReader();
        SepColumnMap columns = ResolveColumns(reader);
        return reader.ParallelEnumerate(
            row => ParseSepRow(row, columns),
            degreeOfParallelism).ToArray();
    }

    internal void ValidateParallel(CsvBenchmarkRow[] officeRows, CsvBenchmarkRow[] sepRows) {
        Validate(nameof(OfficeIMOParallel), officeRows);
        Validate(nameof(SepParallel), sepRows);
    }

    internal void ValidateOfficeParallel(CsvBenchmarkRow[] officeRows) =>
        Validate(nameof(OfficeIMORecordParallel), officeRows);

    internal void ValidateSequential(
        CsvBenchmarkRow[] rowsAsRows,
        CsvBenchmarkRow[] manualRows,
        CsvBenchmarkRow[] sepRows) {
        Validate(nameof(OfficeIMORowsAs), rowsAsRows);
        Validate(nameof(OfficeIMOManual), manualRows);
        Validate(nameof(SepSequential), sepRows);
    }

    private SepReader CreateSepReader() =>
        SepLib.New(',').Reader(value => value with { Unescape = true }).FromText(_csvText);

    private static SepColumnMap ResolveColumns(SepReader reader) => new(
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Id)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Name)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Department)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Region)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.IsEnabled)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Created)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Score)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Owner)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.TicketCount)),
        reader.Header.IndexOf(nameof(CsvBenchmarkRow.Notes)));

    private static CsvBenchmarkRow ParseSepRow(SepReader.Row row, SepColumnMap columns) => new() {
        Id = row[columns.Id].Parse<int>(),
        Name = row[columns.Name].ToString(),
        Department = row[columns.Department].ToString(),
        Region = row[columns.Region].ToString(),
        IsEnabled = row[columns.IsEnabled].Parse<bool>(),
        Created = DateTime.Parse(
            row[columns.Created].Span,
            Invariant,
            DateTimeStyles.RoundtripKind),
        Score = row[columns.Score].Parse<decimal>(),
        Owner = row[columns.Owner].ToString(),
        TicketCount = row[columns.TicketCount].Parse<int>(),
        Notes = row[columns.Notes].ToString()
    };

    private void Validate(string library, CsvBenchmarkRow[] actualRows) {
        if (actualRows.Length != _expectedRows.Length) {
            throw new InvalidDataException(
                $"{library} produced {actualRows.Length} rows instead of {_expectedRows.Length}.");
        }

        for (int index = 0; index < actualRows.Length; index++) {
            if (!RowsEqual(_expectedRows[index], actualRows[index])) {
                throw new InvalidDataException($"{library} produced a different typed row at index {index}.");
            }
        }
    }

    private CsvBenchmarkRow[] ReadOfficeImoRows() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(_csvText);
        return reader.RowsAs<CsvBenchmarkRow>().ToArray();
    }

    private CsvBenchmarkRow[] ReadOfficeImoManualRows() {
        using DbDataReader reader = CsvDocument.OpenTextDataReader(_csvText);
        int id = reader.GetOrdinal(nameof(CsvBenchmarkRow.Id));
        int name = reader.GetOrdinal(nameof(CsvBenchmarkRow.Name));
        int department = reader.GetOrdinal(nameof(CsvBenchmarkRow.Department));
        int region = reader.GetOrdinal(nameof(CsvBenchmarkRow.Region));
        int isEnabled = reader.GetOrdinal(nameof(CsvBenchmarkRow.IsEnabled));
        int created = reader.GetOrdinal(nameof(CsvBenchmarkRow.Created));
        int score = reader.GetOrdinal(nameof(CsvBenchmarkRow.Score));
        int owner = reader.GetOrdinal(nameof(CsvBenchmarkRow.Owner));
        int ticketCount = reader.GetOrdinal(nameof(CsvBenchmarkRow.TicketCount));
        int notes = reader.GetOrdinal(nameof(CsvBenchmarkRow.Notes));
        var rows = new List<CsvBenchmarkRow>(RowCount);
        while (reader.Read()) {
            rows.Add(new CsvBenchmarkRow {
                Id = reader.GetInt32(id),
                Name = reader.GetString(name),
                Department = reader.GetString(department),
                Region = reader.GetString(region),
                IsEnabled = reader.GetBoolean(isEnabled),
                Created = reader.GetDateTime(created),
                Score = reader.GetDecimal(score),
                Owner = reader.GetString(owner),
                TicketCount = reader.GetInt32(ticketCount),
                Notes = reader.GetString(notes)
            });
        }

        return rows.ToArray();
    }

    private static bool RowsEqual(CsvBenchmarkRow expected, CsvBenchmarkRow actual) =>
        expected.Id == actual.Id
        && expected.Name == actual.Name
        && expected.Department == actual.Department
        && expected.Region == actual.Region
        && expected.IsEnabled == actual.IsEnabled
        && expected.Created == actual.Created
        && expected.Created.Kind == actual.Created.Kind
        && expected.Score == actual.Score
        && expected.Owner == actual.Owner
        && expected.TicketCount == actual.TicketCount
        && expected.Notes == actual.Notes;

    private readonly record struct SepColumnMap(
        int Id,
        int Name,
        int Department,
        int Region,
        int IsEnabled,
        int Created,
        int Score,
        int Owner,
        int TicketCount,
        int Notes);
}

internal struct CsvObservationAccumulator {
    private const ulong ChecksumOffset = 14695981039346656037UL;
    private const ulong ChecksumPrime = 1099511628211UL;
    private int _cells;
    private long _characters;
    private ulong _checksum;

    internal void Add(string? value) => Add(value.AsSpan());

    internal void Add(ReadOnlySpan<char> value) {
        if (_cells == 0) {
            _checksum = ChecksumOffset;
        }

        _cells++;
        _characters += value.Length;
        foreach (char character in value) {
            _checksum ^= character;
            _checksum *= ChecksumPrime;
        }

        _checksum ^= (ulong)value.Length;
        _checksum *= ChecksumPrime;
    }

    internal CsvReadObservation ToObservation(int rows) =>
        new(rows, _cells, _characters, _cells == 0 ? ChecksumOffset : _checksum);
}
