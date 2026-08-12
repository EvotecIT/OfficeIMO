#nullable enable

using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using Sylvan.Data.Csv;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Compares database-shaped CSV exports across ordinary, quoted, multiline,
/// and nullable values. This complements the numeric-heavy 40-column lane.
/// </summary>
[MemoryDiagnoser]
public class CsvDataReaderWriteBenchmarks
{
    private static readonly string[] Headers =
    [
        "Id", "Name", "Department", "Region", "IsEnabled",
        "Created", "Score", "Owner", "TicketCount", "Notes"
    ];

    private static readonly Type[] FieldTypes =
    [
        typeof(int), typeof(string), typeof(string), typeof(string), typeof(bool),
        typeof(DateTime), typeof(decimal), typeof(object), typeof(int), typeof(object)
    ];

    private static readonly CsvDataWriterOptions SylvanWriterOptions = new() { NewLine = "\n" };

    private object?[][] _rows = [];
    private bool _captureOutput;
    private string? _capturedOutput;

    public int ParallelDegree { get; set; } = 4;

    public int ParallelBatchSize { get; set; } = 4096;

    [Params(25000, 100000)]
    public int RowCount { get; set; }

    [Params(CsvBenchmarkShape.Mixed, CsvBenchmarkShape.Quoted, CsvBenchmarkShape.Multiline)]
    public CsvBenchmarkShape Shape { get; set; }

    public void Setup()
    {
        Initialize();
        ValidateOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader);
        ValidateOutput(nameof(OfficeIMO_WriteDataReaderParallel), OfficeIMO_WriteDataReaderParallel);
        ValidateOutput(nameof(Sylvan_WriteDataReader), Sylvan_WriteDataReader);
    }

    public void SetupOfficeIMOAndSylvan()
    {
        Initialize();
        ValidateOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader);
        ValidateOutput(nameof(Sylvan_WriteDataReader), Sylvan_WriteDataReader);
    }

    public void SetupOfficeIMOSequentialAndParallel()
    {
        Initialize();
        ValidateOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader);
        ValidateOutput(nameof(OfficeIMO_WriteDataReaderParallel), OfficeIMO_WriteDataReaderParallel);
    }

    [GlobalSetup(Target = nameof(OfficeIMO_WriteDataReader))]
    public void SetupOfficeIMO()
    {
        Initialize();
        ValidateOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader);
    }

    [GlobalSetup(Target = nameof(Sylvan_WriteDataReader))]
    public void SetupSylvan()
    {
        Initialize();
        ValidateOutput(nameof(Sylvan_WriteDataReader), Sylvan_WriteDataReader);
    }

    [GlobalSetup(Target = nameof(OfficeIMO_WriteDataReaderParallel))]
    public void SetupOfficeIMOParallel()
    {
        Initialize();
        ValidateOutput(nameof(OfficeIMO_WriteDataReaderParallel), OfficeIMO_WriteDataReaderParallel);
    }

    private void Initialize()
    {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority))
        {
            BenchmarkProcessorAffinity.ApplyPriority(priority);
        }

        var source = CsvBenchmarkData.Create(RowCount, Shape);
        _rows = new object?[source.Length][];
        for (var i = 0; i < source.Length; i++)
        {
            var item = source[i];
            _rows[i] =
            [
                item.Id,
                item.Name,
                item.Department,
                item.Region,
                item.IsEnabled,
                item.Created,
                item.Score,
                i % 19 == 0 ? DBNull.Value : item.Owner,
                item.TicketCount,
                i % 23 == 0 ? DBNull.Value : item.Notes
            ];
        }

    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_WriteDataReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, FieldTypes);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n" });
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteDataReaderParallel()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, FieldTypes);
        CsvDocument.WriteDataReaderParallel(
            writer,
            reader,
            new CsvSaveOptions { NewLine = "\n" },
            new CsvWriteParallelOptions
            {
                MaxDegreeOfParallelism = ParallelDegree,
                BatchSize = ParallelBatchSize
            });
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sylvan_WriteDataReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows, FieldTypes);
        using var csv = CsvDataWriter.Create(writer, SylvanWriterOptions);
        csv.Write(reader);
        return CompleteWrite(writer);
    }

    private void ValidateOutput(string method, Func<int> write)
    {
        _captureOutput = true;
        _capturedOutput = null;
        try
        {
            var reportedLength = write();
            var output = _capturedOutput
                ?? throw new InvalidOperationException($"{method} did not expose its output to benchmark preflight.");
            if (reportedLength != output.Length)
            {
                throw new InvalidOperationException($"{method} reported {reportedLength} characters but produced {output.Length}.");
            }

            CsvBenchmarkOutputValidator.Validate(
                method,
                output,
                Headers,
                RowCount,
                expectedTextRows: null,
                expectedObjectRows: _rows);
        }
        finally
        {
            _captureOutput = false;
            _capturedOutput = null;
        }
    }

    private int CompleteWrite(StringWriter writer)
    {
        var output = writer.GetStringBuilder();
        if (_captureOutput)
        {
            _capturedOutput = output.ToString();
        }

        return output.Length;
    }
}
