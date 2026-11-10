using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using OfficeIMO.Data;
using RecordParser.Builders.Reader;
using RecordParser.Extensions;
using RecordParser.Parsers;
using SepLib = nietras.SeparatedValues.Sep;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Compares ordered three-column projections from the shared quoted ten-column input.</summary>
[MemoryDiagnoser]
public class CsvProjectedParallelBenchmarks {
    private const int RowCount = 100_000;
    private string _text = string.Empty;
    private CsvBenchmarkRow[] _expected = [];
    private ProjectedRow[]? _last;
    private static readonly IVariableLengthReader<ProjectedRow> RecordParserReader =
        new VariableLengthReaderBuilder<ProjectedRow>()
            .Map(row => row.Id, 0)
            .Map(row => row.Score, 6)
            .Map(row => row.TicketCount, 8)
            .Build(",", CultureInfo.InvariantCulture);

    [Params(1, 4, 8, 16)]
    public int DegreeOfParallelism { get; set; }

    [GlobalSetup]
    public void Setup() {
        global::OfficeIMO.Benchmarks.BenchmarkProcessPriority.ApplyConfiguredPriority();
        _expected = CsvBenchmarkData.Create(RowCount, CsvBenchmarkShape.Quoted);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, _expected,
            new CsvSaveOptions { NewLine = "\n", DateTimeFormat = "O" });
        _text = writer.ToString();
        Validate(OfficeIMO());
        Validate(Sep());
        Validate(RecordParser());
        Console.WriteLine("PROJECTED100K10_INPUT_UTF8_SHA256=" +
            Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(_text))));
    }

    [Benchmark]
    public ProjectedRow[] OfficeIMO() => _last = CsvDocument.ReadTextRowsAsParallel<ProjectedRow>(
        _text,
        header => {
            foreach (string name in CsvBenchmarkData.Headers) header.GetOrdinal(name);
            int id = header.GetOrdinal(nameof(CsvBenchmarkRow.Id));
            int score = header.GetOrdinal(nameof(CsvBenchmarkRow.Score));
            int tickets = header.GetOrdinal(nameof(CsvBenchmarkRow.TicketCount));
            return row => new ProjectedRow {
                Id = row.GetInt32(id),
                Score = row.GetDecimal(score),
                TicketCount = row.GetInt32(tickets)
            };
        },
        options: new CsvReaderOptions { Culture = CultureInfo.InvariantCulture },
        parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = DegreeOfParallelism }).ToArray();

    [Benchmark]
    public ProjectedRow[] Sep() {
        using var reader = SepLib.New(',').Reader(options => options with { Unescape = true }).FromText(_text);
        if (!reader.Header.ColNames.SequenceEqual(CsvBenchmarkData.Headers))
            throw new InvalidDataException("Unexpected projection header.");
        int id = reader.Header.IndexOf(nameof(CsvBenchmarkRow.Id));
        int score = reader.Header.IndexOf(nameof(CsvBenchmarkRow.Score));
        int tickets = reader.Header.IndexOf(nameof(CsvBenchmarkRow.TicketCount));
        return _last = DegreeOfParallelism == 1
            ? reader.Enumerate(row => ParseSep(row, id, score, tickets)).ToArray()
            : reader.ParallelEnumerate(row => ParseSep(row, id, score, tickets), DegreeOfParallelism).ToArray();
    }

    [Benchmark]
    public ProjectedRow[] RecordParser() {
        using var reader = new StringReader(_text);
        if (reader.ReadLine() != string.Join(",", CsvBenchmarkData.Headers))
            throw new InvalidDataException("Unexpected projection header.");
        return _last = reader.ReadRecords(RecordParserReader, new VariableLengthReaderOptions {
            HasHeader = false,
            ContainsQuotedFields = true,
            ParallelismOptions = new() {
                Enabled = DegreeOfParallelism > 1,
                EnsureOriginalOrdering = true,
                MaxDegreeOfParallelism = DegreeOfParallelism
            }
        }).ToArray();
    }

    private static ProjectedRow ParseSep(SepReader.Row row, int id, int score, int tickets) => new() {
        Id = row[id].Parse<int>(),
        Score = row[score].Parse<decimal>(),
        TicketCount = row[tickets].Parse<int>()
    };

    [GlobalCleanup]
    public void ValidateFinal() {
        Validate(_last ?? throw new InvalidDataException("Missing projection output."));
        Console.WriteLine("FINAL_100000x3_FIELDS_ORDER_COUNT_ORACLE_PASS");
    }

    private void Validate(ProjectedRow[] actual) {
        if (actual.Length != RowCount) throw new InvalidDataException("Incomplete projection.");
        for (int index = 0; index < actual.Length; index++) {
            CsvBenchmarkRow expected = _expected[index];
            ProjectedRow row = actual[index];
            if (row.Id != expected.Id || row.Score != expected.Score || row.TicketCount != expected.TicketCount)
                throw new InvalidDataException($"Incorrect projected row {index}.");
        }
    }

    /// <summary>Every lane owns the same three typed values for each input record.</summary>
    public sealed class ProjectedRow {
        public int Id { get; set; }
        public decimal Score { get; set; }
        public int TicketCount { get; set; }
    }
}
