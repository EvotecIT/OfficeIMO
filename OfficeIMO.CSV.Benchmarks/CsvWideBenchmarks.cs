#nullable enable

using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using nietras.SeparatedValues;
using CsvHelperConfiguration = CsvHelper.Configuration.CsvConfiguration;
using CsvHelperReader = CsvHelper.CsvReader;
using CsvHelperWriter = CsvHelper.CsvWriter;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using DataplatCsvReaderOptions = Dataplat.Dbatools.Csv.Reader.CsvReaderOptions;
using DataplatCsvWriter = Dataplat.Dbatools.Csv.Writer.CsvWriter;
using DataplatCsvWriterOptions = Dataplat.Dbatools.Csv.Writer.CsvWriterOptions;
using SepLib = nietras.SeparatedValues.Sep;
using SepReaderOptions = nietras.SeparatedValues.SepReaderOptions;
using SepWriterOptions = nietras.SeparatedValues.SepWriterOptions;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;
using SylvanCsvDataWriter = Sylvan.Data.Csv.CsvDataWriter;
using SylvanCsvDataWriterOptions = Sylvan.Data.Csv.CsvDataWriterOptions;

namespace OfficeIMO.CSV.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class CsvWideBenchmarks
{
    private static readonly string[] Headers = CreateHeaders();
    private static readonly DataplatCsvReaderOptions DataplatReaderOptions = new() { HasHeaderRow = true };
    private static readonly DataplatCsvWriterOptions DataplatWriterOptions = new() { NewLine = "\n" };
    private static readonly CsvHelperConfiguration CsvHelperWriteConfiguration = new(CultureInfo.InvariantCulture) { NewLine = "\n" };
    private static readonly SepReaderOptions SepReadOptions = SepLib.New(',').Reader(options => options with { Unescape = true });
    private static readonly SepWriterOptions SepWriteOptions = SepLib.New(',').Writer(options => options with { WriteHeader = true, Escape = true });
    private static readonly SylvanCsvDataWriterOptions SylvanWriterOptions = new() { NewLine = "\n" };

    private object?[][] _rows = [];
    private string?[][] _textRows = [];
    private string _csvText = string.Empty;
    private bool _captureWriteOutput;
    private string? _capturedWriteOutput;
    private string _csvPath = string.Empty;
    private CsvSchema _wideSchema = new CsvSchemaBuilder().Build();

    [Params(1000, 10000, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _rows = CsvWideBenchmarkData.Create(RowCount);
        _textRows = _rows.Select(ProjectTextRow).ToArray();

        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csv.WriteRows(Headers, _rows);

        _csvText = writer.ToString();
        _csvPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO.CSV.Benchmarks.{Guid.NewGuid():N}.csv");
        File.WriteAllText(_csvPath, _csvText, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

        var schema = new CsvSchemaBuilder();
        schema.Column(Headers[0]).AsInt32();
        schema.Column(Headers[1]).AsString();
        schema.Column(Headers[2]).AsDateTime();
        schema.Column(Headers[3]).AsBoolean();
        for (var i = 4; i < Headers.Length; i++)
        {
            schema.Column(Headers[i]).AsType(typeof(decimal));
        }

        _wideSchema = schema.Build();
        ValidateWriteBenchmarkOutputs();
    }

    private void ValidateWriteBenchmarkOutputs()
    {
        ValidateWriteOutput(nameof(OfficeIMO_WriteProjectedRows), OfficeIMO_WriteProjectedRows, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteTrustedProjectedRows), OfficeIMO_WriteTrustedProjectedRows, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(CsvHelper_WriteProjectedRows), CsvHelper_WriteProjectedRows, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(Sylvan_WriteProjectedRows), Sylvan_WriteProjectedRows, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(Dataplat_WriteProjectedRows), Dataplat_WriteProjectedRows, expectedObjectRows: _rows);
        ValidateWriteOutput(nameof(Dataplat_WriteFromReader), Dataplat_WriteFromReader, expectedObjectRows: _rows);

        ValidateWriteOutput(nameof(OfficeIMO_WriteValidatedTextRows), OfficeIMO_WriteValidatedTextRows, _textRows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteTrustedTextRows), OfficeIMO_WriteTrustedTextRows, _textRows);
        ValidateWriteOutput(nameof(CsvHelper_WriteTextRows), CsvHelper_WriteTextRows, _textRows);
        ValidateWriteOutput(nameof(Sylvan_WriteTextRows), Sylvan_WriteTextRows, _textRows);
        ValidateWriteOutput(nameof(Dataplat_WriteTextRows), Dataplat_WriteTextRows, _textRows);
        ValidateWriteOutput(nameof(Sep_WriteProjectedRows), Sep_WriteProjectedRows, _textRows);
    }

    private void ValidateWriteOutput(
        string method,
        Func<int> write,
        string?[][]? expectedTextRows = null,
        object?[][]? expectedObjectRows = null)
    {
        _captureWriteOutput = true;
        _capturedWriteOutput = null;
        try
        {
            var reportedLength = write();
            var output = _capturedWriteOutput
                ?? throw new InvalidOperationException($"{method} did not expose its output to benchmark preflight.");
            if (reportedLength != output.Length)
            {
                throw new InvalidOperationException($"{method} reported {reportedLength} characters but produced {output.Length}.");
            }

            CsvBenchmarkOutputValidator.Validate(method, output, Headers, RowCount, expectedTextRows, expectedObjectRows);
        }
        finally
        {
            _captureWriteOutput = false;
            _capturedWriteOutput = null;
        }
    }

    private int CompleteWrite(StringWriter writer)
    {
        var buffer = writer.GetStringBuilder();
        if (_captureWriteOutput)
        {
            _capturedWriteOutput = buffer.ToString();
        }

        return buffer.Length;
    }

    private int CompleteWrite(string output)
    {
        if (_captureWriteOutput)
        {
            _capturedWriteOutput = output;
        }

        return output.Length;
    }

    [GlobalCleanup]
    public void Cleanup()
    {
        if (_csvPath.Length > 0 && File.Exists(_csvPath))
        {
            File.Delete(_csvPath);
        }
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csv.WriteRows(Headers, _rows);

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteTrustedProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        if (_rows.Length == 0)
        {
            return 0;
        }

        csv.WriteRow(Headers, _rows[0]);
        for (var i = 1; i < _rows.Length; i++)
        {
            csv.WriteTrustedRow(_rows[i]);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteValidatedTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csv.WriteTextRows(Headers, _textRows);

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteTrustedTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        if (_textRows.Length == 0)
        {
            return 0;
        }

        csv.WriteRow(Headers, _textRows[0]);
        for (var i = 1; i < _textRows.Length; i++)
        {
            csv.WriteTrustedTextRow(_textRows[i]);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteDataReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n" });
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int CsvHelper_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CsvHelperWriteConfiguration);
        foreach (var header in Headers)
        {
            csv.WriteField(header);
        }

        csv.NextRecord();
        foreach (var row in _rows)
        {
            foreach (var value in row)
            {
                csv.WriteField(value);
            }

            csv.NextRecord();
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int CsvHelper_WriteTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CsvHelperWriteConfiguration);
        foreach (var header in Headers)
        {
            csv.WriteField(header);
        }

        csv.NextRecord();
        foreach (string?[] row in _textRows)
        {
            foreach (string? value in row)
            {
                csv.WriteField(value);
            }

            csv.NextRecord();
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sylvan_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows);
        using var csv = SylvanCsvDataWriter.Create(writer, SylvanWriterOptions);
        csv.Write(reader);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sylvan_WriteTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _textRows);
        using var csv = SylvanCsvDataWriter.Create(writer, SylvanWriterOptions);
        csv.Write(reader);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Dataplat_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new DataplatCsvWriter(writer, DataplatWriterOptions);
        csv.WriteHeader(Headers);
        foreach (var row in _rows)
        {
            csv.WriteRow(row);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Dataplat_WriteTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new DataplatCsvWriter(writer, DataplatWriterOptions);
        csv.WriteHeader(Headers);
        foreach (string?[] row in _textRows)
        {
            csv.WriteRow(row);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Dataplat_WriteFromReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows);
        using var csv = new DataplatCsvWriter(writer, DataplatWriterOptions);
        csv.WriteFromReader(reader);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sep_WriteProjectedRows()
    {
        var options = SepWriteOptions;
        using var csv = options.ToText();
        foreach (string?[] row in _textRows)
        {
            using var csvRow = csv.NewRow();
            for (var i = 0; i < Headers.Length; i++)
            {
                csvRow[Headers[i]].Set(row[i].AsSpan());
            }
        }

        return CompleteWrite(csv.ToString());
    }

    [Benchmark]
    public int OfficeIMO_ReadRowsReusableCallback()
    {
        using var reader = new StringReader(_csvText);
        var checksum = 0;
        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            checksum += MeasureValues(values);
        });

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadRowFieldSpansMaterialized()
    {
        using var reader = new StringReader(_csvText);
        var visitor = new CsvMaterializingRowFieldSpanVisitor();
        CsvDocument.ReadRowFieldSpans(reader, ref visitor);
        return visitor.FieldCount + visitor.TextLength;
    }

    [Benchmark]
    public int OfficeIMO_ReadTextRowFieldSpansMaterialized()
    {
        var visitor = new CsvMaterializingRowFieldSpanVisitor();
        CsvDocument.ReadRowFieldSpansFromText(_csvText, ref visitor);
        return visitor.FieldCount + visitor.TextLength;
    }

    [Benchmark]
    public int OfficeIMO_ReadRecordsReusableSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var checksum = 0;
        CsvDocument.ReadRecordsReusable(
            reader,
            values =>
            {
                checksum += MeasureValues(values);
            },
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadFieldSpansSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var checksum = 0;
        CsvDocument.ReadFieldSpans(
            reader,
            (_, _, value) =>
            {
                checksum += 1 + value.Length;
            },
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadFieldSpanVisitorSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var visitor = new CountingFieldSpanVisitor();
        CsvDocument.ReadFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return visitor.FieldCount + visitor.TextLength;
    }

    [Benchmark]
    public int OfficeIMO_ReadTextFieldSpanVisitorSkipHeader()
    {
        var visitor = new CountingFieldSpanVisitor();
        CsvDocument.ReadFieldSpansFromText(
            _csvText,
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return visitor.FieldCount + visitor.TextLength;
    }

    [Benchmark]
    public int OfficeIMO_ReadDataTableStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        var table = document.ToDataTable();
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataTableInferredSchema()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var table = document.ToDataTable(new CsvDataTableOptions { InferSchema = true, SchemaSampleSize = RowCount });
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataTableLoadDataReaderStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader();
        var table = new System.Data.DataTable();
        table.Load(csv);
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataTableLoadDataReaderInferredSchema()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = RowCount });
        var table = new System.Data.DataTable();
        table.Load(csv);
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataReaderStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader();
        return DataTableBenchmarkUtilities.Measure(csv);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataReaderGetStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader();
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                fieldCount += 1 + csv.GetString(i).Length;
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int OfficeIMO_ReadDataReaderInferredSchema()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader(new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = RowCount });
        return DataTableBenchmarkUtilities.Measure(csv);
    }

    [Benchmark]
    public int OfficeIMO_ReadDataReaderExplicitSchema()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        using var csv = document.CreateDataReader(new CsvDataReaderOptions { Schema = _wideSchema });
        return DataTableBenchmarkUtilities.Measure(csv);
    }

    [Benchmark]
    public int OfficeIMO_ReadFileDataReaderExplicitSchema()
    {
        using var csv = CsvDocument.CreateDataReader(
            _csvPath,
            new CsvLoadOptions { Mode = CsvLoadMode.Stream },
            new CsvDataReaderOptions { Schema = _wideSchema });
        return DataTableBenchmarkUtilities.Measure(csv);
    }

    [Benchmark]
    public int OfficeIMO_ReadFieldSpansMaterializedSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var visitor = new CsvMaterializingFieldSpanVisitor();
        CsvDocument.ReadFieldSpans(
            reader,
            ref visitor,
            new CsvLoadOptions { SkipInitialRecords = 1 });
        visitor.Complete();
        return visitor.FieldCount + visitor.TextLength;
    }

    [Benchmark]
    public int CsvHelper_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new CsvHelperReader(reader, CultureInfo.InvariantCulture);
        var fieldCount = 0;
        if (!csv.Read())
        {
            return fieldCount;
        }

        csv.ReadHeader();
        while (csv.Read())
        {
            for (var i = 0; i < Headers.Length; i++)
            {
                string? value = csv.GetField(i);
                fieldCount += 1 + (value?.Length ?? 0);
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int Sylvan_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = SylvanCsvDataReader.Create(reader);
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                fieldCount += 1 + csv.GetString(i).Length;
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int Sylvan_ReadDataTableLoad()
    {
        using var reader = new StringReader(_csvText);
        using var csv = SylvanCsvDataReader.Create(reader);
        var table = new System.Data.DataTable();
        table.Load(csv);
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int Sylvan_ReadFieldSpans()
    {
        using var reader = new StringReader(_csvText);
        using var csv = SylvanCsvDataReader.Create(reader);
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                fieldCount += 1 + csv.GetFieldSpan(i).Length;
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int Dataplat_ReadFields()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new DataplatCsvDataReader(reader, DataplatReaderOptions);
        var fieldCount = 0;
        while (csv.Read())
        {
            for (var i = 0; i < csv.FieldCount; i++)
            {
                fieldCount += 1 + csv.GetString(i).Length;
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int Dataplat_ReadDataTableLoad()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new DataplatCsvDataReader(reader, DataplatReaderOptions);
        var table = new System.Data.DataTable();
        table.Load(csv);
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int Sep_ReadFields()
    {
        var options = SepReadOptions;
        using var csv = options.FromText(_csvText);
        var fieldCount = 0;
        foreach (var row in csv)
        {
            for (var i = 0; i < row.ColCount; i++)
            {
                fieldCount += 1 + row[i].ToString().Length;
            }
        }

        return fieldCount;
    }

    [Benchmark]
    public int Sep_ReadFieldSpans()
    {
        var options = SepReadOptions;
        using var csv = options.FromText(_csvText);
        var fieldCount = 0;
        foreach (var row in csv)
        {
            for (var i = 0; i < row.ColCount; i++)
            {
                fieldCount += 1 + row[i].Span.Length;
            }
        }

        return fieldCount;
    }

    private static string[] CreateHeaders()
    {
        var headers = new string[40];
        headers[0] = "Id";
        headers[1] = "Name";
        headers[2] = "Created";
        headers[3] = "Enabled";
        for (var i = 4; i < headers.Length; i++)
        {
            headers[i] = string.Create(CultureInfo.InvariantCulture, $"Metric{i - 3}");
        }

        return headers;
    }

    private static string?[] ProjectTextRow(object?[] row)
    {
        var values = new string?[row.Length];
        for (var i = 0; i < row.Length; i++)
        {
            values[i] = Convert.ToString(row[i], CultureInfo.InvariantCulture);
        }

        return values;
    }

    private static int MeasureValues(IReadOnlyList<string> values)
    {
        var checksum = 0;
        for (var i = 0; i < values.Count; i++)
        {
            checksum += 1 + values[i].Length;
        }

        return checksum;
    }

    private struct CountingFieldSpanVisitor : ICsvFieldSpanVisitor
    {
        public int FieldCount { get; private set; }

        public int TextLength { get; private set; }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            FieldCount++;
            TextLength += value.Length;
        }

        public bool TryVisitEscapedField(int recordIndex, int fieldIndex, ReadOnlySpan<char> escapedValue, int unescapedLength)
        {
            FieldCount++;
            TextLength += unescapedLength;
            return true;
        }
    }
}

internal static class CsvWideBenchmarkData
{
    public static object?[][] Create(int count)
    {
        var rows = new object?[count][];
        for (var i = 1; i <= count; i++)
        {
            var row = new object?[40];
            row[0] = i;
            row[1] = string.Create(CultureInfo.InvariantCulture, $"Wide-{i:000000}");
            row[2] = new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddSeconds(i);
            row[3] = i % 2 == 0;
            for (var column = 4; column < row.Length; column++)
            {
                row[column] = Math.Round((decimal)(((i + column - 3) * 1.017) % 10000), 4);
            }

            rows[i - 1] = row;
        }

        return rows;
    }
}
