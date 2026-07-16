#nullable enable

using System.Globalization;
using BenchmarkDotNet.Attributes;
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

    [Params(25000)]
    public int RowCount { get; set; }

    [Params(CsvBenchmarkShape.Mixed, CsvBenchmarkShape.Quoted, CsvBenchmarkShape.Multiline)]
    public CsvBenchmarkShape Shape { get; set; }

    [GlobalSetup]
    public void Setup()
    {
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

        ValidateOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader);
        ValidateOutput(nameof(Sylvan_WriteDataReader), Sylvan_WriteDataReader);
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
