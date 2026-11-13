using System.Data;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// Shared shapes, sinks, OfficeIMO lanes and validation of the CSV write comparisons. Every library is
/// configured for the same newline, date format and invariant numbers; setup requires every lane to write
/// output byte-identical to OfficeIMO's data-reader lane and checks that output value by value against the
/// source. Lanes are named Library_SourceSink so each compares with OfficeIMO's matching lane.
/// </summary>
public abstract class CsvWriteBenchmarkBase {
    protected const string DateFormat = CsvWriteSource.DateFormat;
    /// <summary>Sep always ends rows with Environment.NewLine, so every library uses it.</summary>
    protected static readonly string NewLine = Environment.NewLine;
    protected static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private MemoryStream _stream = null!;
    private int _textCapacity;
    private bool _capture;
    private byte[]? _captured;

    /// <summary>Only the date shape asks for a date format, as an export without dates would.</summary>
    protected CsvSaveOptions OfficeOptions { get; private set; } = null!;

    internal CsvWriteSource Source { get; private set; } = null!;

    [Params(CsvWriteShape.QuotedText, CsvWriteShape.Numbers, CsvWriteShape.Dates, CsvWriteShape.Wide)]
    public CsvWriteShape Shape { get; set; }

    [GlobalSetup]
    public void Setup() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        Source = CsvWriteSource.Create(Shape);
        OfficeOptions = new CsvSaveOptions { NewLine = NewLine, DateTimeFormat = Shape == CsvWriteShape.Dates ? DateFormat : null };
        PrepareCompetitors();
        _stream = new MemoryStream();
        byte[] reference = Capture(nameof(OfficeIMO_DataReaderUtf8), OfficeIMO_DataReaderUtf8);
        CsvBenchmarkOutputValidator.Validate(nameof(OfficeIMO_DataReaderUtf8), Utf8NoBom.GetString(reference),
            Source.Headers, Source.Rows.Length, expectedTextRows: null, expectedObjectRows: Source.Rows);
        _stream = new MemoryStream(reference.Length);
        _textCapacity = Utf8NoBom.GetCharCount(reference);
        IEnumerable<(string Name, Func<int> Write)> office = [
            (nameof(OfficeIMO_ObjectsUtf8), OfficeIMO_ObjectsUtf8), (nameof(OfficeIMO_ObjectsText), OfficeIMO_ObjectsText),
            (nameof(OfficeIMO_DataTableUtf8), OfficeIMO_DataTableUtf8), (nameof(OfficeIMO_DataTableText), OfficeIMO_DataTableText),
            (nameof(OfficeIMO_DataReaderText), OfficeIMO_DataReaderText)];
        foreach (var (name, write) in office.Concat(CompetitorLanes())) {
            byte[] output = Capture(name, write);
            if (!output.AsSpan().SequenceEqual(reference)) {
                int index = output.AsSpan().CommonPrefixLength(reference);
                string Around(byte[] bytes) => Utf8NoBom.GetString(bytes, Math.Max(0, index - 40), Math.Min(80, bytes.Length - Math.Max(0, index - 40)))
                    .Replace("\r", "\\r").Replace("\n", "\\n");
                throw new InvalidOperationException($"{name} ({Shape}) differs from the reference at byte {index}: '{Around(output)}' vs '{Around(reference)}'.");
            }
        }
    }

    /// <summary>Prepares library-specific inputs outside the measurement.</summary>
    protected virtual void PrepareCompetitors() { }

    /// <summary>The other libraries' lanes, each checked against the reference output.</summary>
    protected abstract IEnumerable<(string Name, Func<int> Write)> CompetitorLanes();

    [Benchmark]
    public int OfficeIMO_ObjectsUtf8() {
        using (CsvRowWriter csv = CsvRowWriter.CreateStream(ResetStream(), OfficeOptions, leaveOpen: true)) {
            foreach (ICsvWriteRecord record in Source.Records) csv.WriteObject(record);
        }
        return CompleteUtf8();
    }

    [Benchmark]
    public int OfficeIMO_ObjectsText() {
        StringWriter writer = NewTextWriter();
        using (var csv = new CsvRowWriter(writer, OfficeOptions, leaveOpen: true)) {
            foreach (ICsvWriteRecord record in Source.Records) csv.WriteObject(record);
        }
        return CompleteText(writer);
    }

    /// <summary>OfficeIMO takes the table itself; the other libraries read it through its data reader, which is how they accept a table.</summary>
    [Benchmark]
    public int OfficeIMO_DataTableUtf8() {
        CsvDocument.WriteDataTable(ResetStream(), Source.Table, OfficeOptions, leaveOpen: true);
        return CompleteUtf8();
    }

    [Benchmark]
    public int OfficeIMO_DataTableText() {
        StringWriter writer = NewTextWriter();
        CsvDocument.WriteDataTable(writer, Source.Table, OfficeOptions);
        return CompleteText(writer);
    }

    [Benchmark]
    public int OfficeIMO_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        CsvDocument.WriteDataReader(ResetStream(), reader, OfficeOptions, leaveOpen: true);
        return CompleteUtf8();
    }

    [Benchmark]
    public int OfficeIMO_DataReaderText() {
        StringWriter writer = NewTextWriter();
        using BenchmarkArrayDataReader reader = NewDataReader();
        CsvDocument.WriteDataReader(writer, reader, OfficeOptions);
        return CompleteText(writer);
    }

    // Sinks: one reusable stream and a text buffer sized for the output, the same for every library.

    internal BenchmarkArrayDataReader NewDataReader() => new(Source.Headers, Source.Rows, Source.Types);

    protected MemoryStream ResetStream() {
        _stream.Position = 0;
        _stream.SetLength(0);
        return _stream;
    }

    protected StringWriter NewTextWriter() => new(new StringBuilder(_textCapacity), CultureInfo.InvariantCulture);

    protected int CompleteUtf8() {
        if (_capture) _captured = _stream.ToArray();
        return (int)_stream.Length;
    }

    protected int CompleteText(StringWriter writer) {
        StringBuilder text = writer.GetStringBuilder();
        if (_capture) _captured = Utf8NoBom.GetBytes(text.ToString());
        return text.Length;
    }

    private byte[] Capture(string name, Func<int> write) {
        _capture = true;
        _captured = null;
        try {
            int length = write();
            byte[] output = _captured ?? throw new InvalidOperationException($"{name} did not expose its output.");
            if (output.Length == 0 || length <= 0) throw new InvalidOperationException($"{name} wrote nothing.");
            return output;
        } finally {
            _capture = false;
        }
    }
}
