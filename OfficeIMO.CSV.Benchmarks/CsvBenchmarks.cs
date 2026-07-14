#nullable enable

using System.Globalization;
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
public class CsvBenchmarks
{
    private static readonly string[] Headers =
    [
        nameof(CsvBenchmarkRow.Id),
        nameof(CsvBenchmarkRow.Name),
        nameof(CsvBenchmarkRow.Department),
        nameof(CsvBenchmarkRow.Region),
        nameof(CsvBenchmarkRow.IsEnabled),
        nameof(CsvBenchmarkRow.Created),
        nameof(CsvBenchmarkRow.Score),
        nameof(CsvBenchmarkRow.Owner),
        nameof(CsvBenchmarkRow.TicketCount),
        nameof(CsvBenchmarkRow.Notes)
    ];

    private CsvBenchmarkRow[] _rows = [];
    private object?[][] _projectedRows = [];
    private string?[][] _projectedTextRows = [];
    private string _csvText = string.Empty;
    private bool _captureWriteOutput;
    private string? _capturedWriteOutput;
    private static readonly DataplatCsvReaderOptions DataplatReaderOptions = new() { HasHeaderRow = true };
    private static readonly DataplatCsvWriterOptions DataplatWriterOptions = new() { NewLine = "\n" };
    private static readonly CsvHelperConfiguration CsvHelperWriteConfiguration = new(CultureInfo.InvariantCulture) { NewLine = "\n" };
    private static readonly SepReaderOptions SepReadOptions = SepLib.New(',').Reader(options => options with { Unescape = true });
    private static readonly SepWriterOptions SepWriteOptions = SepLib.New(',').Writer(options => options with { WriteHeader = true, Escape = true });
    private static readonly SylvanCsvDataWriterOptions SylvanWriterOptions = new() { NewLine = "\n" };

    [Params(1000, 10000, 25000)]
    public int RowCount { get; set; }

    [Params(CsvBenchmarkShape.Mixed, CsvBenchmarkShape.Quoted, CsvBenchmarkShape.Multiline)]
    public CsvBenchmarkShape Shape { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _rows = CsvBenchmarkData.Create(RowCount, Shape);
        _projectedRows = _rows.Select(ProjectRow).ToArray();
        _projectedTextRows = _projectedRows.Select(ProjectTextRow).ToArray();

        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, _rows, new CsvSaveOptions { NewLine = "\n" });
        _csvText = writer.ToString();

        ValidateWriteBenchmarkOutputs();
    }

    private void ValidateWriteBenchmarkOutputs()
    {
        ValidateWriteOutput(nameof(OfficeIMO_WriteObjects), OfficeIMO_WriteObjects, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteProjectedRows), OfficeIMO_WriteProjectedRows, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteTrustedProjectedRows), OfficeIMO_WriteTrustedProjectedRows, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteDataReader), OfficeIMO_WriteDataReader, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(CsvHelper_WriteTypedRecords), CsvHelper_WriteTypedRecords, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(CsvHelper_WriteProjectedRows), CsvHelper_WriteProjectedRows, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(Sylvan_WriteProjectedRows), Sylvan_WriteProjectedRows, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(Dataplat_WriteProjectedRows), Dataplat_WriteProjectedRows, expectedObjectRows: _projectedRows);
        ValidateWriteOutput(nameof(Dataplat_WriteFromReader), Dataplat_WriteFromReader, expectedObjectRows: _projectedRows);

        ValidateWriteOutput(nameof(OfficeIMO_WriteValidatedTextRows), OfficeIMO_WriteValidatedTextRows, _projectedTextRows);
        ValidateWriteOutput(nameof(OfficeIMO_WriteTrustedTextRows), OfficeIMO_WriteTrustedTextRows, _projectedTextRows);
        ValidateWriteOutput(nameof(CsvHelper_WriteTextRows), CsvHelper_WriteTextRows, _projectedTextRows);
        ValidateWriteOutput(nameof(Sylvan_WriteTextRows), Sylvan_WriteTextRows, _projectedTextRows);
        ValidateWriteOutput(nameof(Dataplat_WriteTextRows), Dataplat_WriteTextRows, _projectedTextRows);
        ValidateWriteOutput(nameof(Sep_WriteProjectedRows), Sep_WriteProjectedRows, _projectedTextRows);
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

    [Benchmark(Baseline = true)]
    public int OfficeIMO_WriteObjects()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, _rows, new CsvSaveOptions { NewLine = "\n" });
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csv.WriteRows(Headers, _projectedRows);

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteTrustedProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        if (_projectedRows.Length == 0)
        {
            return 0;
        }

        csv.WriteRow(Headers, _projectedRows[0]);
        for (var i = 1; i < _projectedRows.Length; i++)
        {
            csv.WriteTrustedRow(_projectedRows[i]);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteValidatedTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        csv.WriteTextRows(Headers, _projectedTextRows);

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteTrustedTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        if (_projectedTextRows.Length == 0)
        {
            return 0;
        }

        csv.WriteRow(Headers, _projectedTextRows[0]);
        for (var i = 1; i < _projectedTextRows.Length; i++)
        {
            csv.WriteTrustedTextRow(_projectedTextRows[i]);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteDataReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _projectedRows);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions { NewLine = "\n" });
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int OfficeIMO_WriteProjectedRowsAlwaysQuoted()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n", QuoteMode = CsvQuoteMode.Always }, leaveOpen: true);
        foreach (object?[] row in _projectedRows)
        {
            csv.WriteRow(Headers, row);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int CsvHelper_WriteTypedRecords()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CsvHelperWriteConfiguration);
        csv.WriteRecords(_rows);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int CsvHelper_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CsvHelperWriteConfiguration);
        foreach (string header in Headers)
        {
            csv.WriteField(header);
        }

        csv.NextRecord();

        foreach (object?[] row in _projectedRows)
        {
            foreach (object? value in row)
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
        foreach (string header in Headers)
        {
            csv.WriteField(header);
        }

        csv.NextRecord();

        foreach (string?[] row in _projectedTextRows)
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
        using var reader = new BenchmarkArrayDataReader(Headers, _projectedRows);
        using var csv = SylvanCsvDataWriter.Create(writer, SylvanWriterOptions);
        csv.Write(reader);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sylvan_WriteTextRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _projectedTextRows);
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
        foreach (object?[] row in _projectedRows)
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
        foreach (string?[] row in _projectedTextRows)
        {
            csv.WriteRow(row);
        }

        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Dataplat_WriteFromReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _projectedRows);
        using var csv = new DataplatCsvWriter(writer, DataplatWriterOptions);
        csv.WriteFromReader(reader);
        return CompleteWrite(writer);
    }

    [Benchmark]
    public int Sep_WriteProjectedRows()
    {
        var options = SepWriteOptions;
        using var csv = options.ToText();
        foreach (string?[] row in _projectedTextRows)
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
    public int OfficeIMO_ReadRowsCallback()
    {
        using var reader = new StringReader(_csvText);
        var checksum = 0;
        CsvDocument.ReadRows(reader, (_, values) =>
        {
            checksum += MeasureValues(values);
        });

        return checksum;
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
    public int OfficeIMO_ReadFieldSpansMaterialized()
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
    public int OfficeIMO_ReadStreamingRows()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var checksum = 0;
        foreach (CsvRow row in document.AsEnumerable())
        {
            checksum += MeasureRow(row);
        }

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadInMemoryRows()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        var checksum = 0;
        foreach (CsvRow row in document.AsEnumerable())
        {
            checksum += MeasureRow(row);
        }

        return checksum;
    }

    [Benchmark]
    public int OfficeIMO_ReadDataTableStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        var table = document.ToDataTable();
        return DataTableBenchmarkUtilities.Measure(table);
    }

    [Benchmark]
    public int OfficeIMO_ReadStreamingDataTableStrings()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
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
    public int OfficeIMO_ReadStreamingDataTableInferredSchemaDefaultSample()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var table = document.ToDataTable(new CsvDataTableOptions { InferSchema = true });
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

    [Benchmark]
    public int CsvHelper_ReadTypedRecords()
    {
        using var reader = new StringReader(_csvText);
        using var csv = new CsvHelperReader(reader, CultureInfo.InvariantCulture);
        var count = 0;
        foreach (CsvBenchmarkRow _ in csv.GetRecords<CsvBenchmarkRow>())
        {
            count++;
        }

        return count;
    }

    private static object?[] ProjectRow(CsvBenchmarkRow row)
    {
        return
        [
            row.Id,
            row.Name,
            row.Department,
            row.Region,
            row.IsEnabled,
            row.Created,
            row.Score,
            row.Owner,
            row.TicketCount,
            row.Notes
        ];
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

    private static int MeasureRow(CsvRow row)
    {
        var checksum = 0;
        for (var i = 0; i < row.FieldCount; i++)
        {
            checksum += 1 + (Convert.ToString(row[i], CultureInfo.InvariantCulture)?.Length ?? 0);
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

public enum CsvBenchmarkShape
{
    Mixed,
    Quoted,
    Multiline
}

public sealed class CsvBenchmarkRow
{
    public int Id { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Department { get; set; } = string.Empty;
    public string Region { get; set; } = string.Empty;
    public bool IsEnabled { get; set; }
    public DateTime Created { get; set; }
    public decimal Score { get; set; }
    public string Owner { get; set; } = string.Empty;
    public int TicketCount { get; set; }
    public string Notes { get; set; } = string.Empty;
}

internal static class CsvBenchmarkData
{
    private static readonly string[] Regions = ["NA", "EU", "APAC", "LATAM"];

    public static CsvBenchmarkRow[] Create(int count, CsvBenchmarkShape shape)
    {
        var rows = new CsvBenchmarkRow[count];
        for (var i = 1; i <= count; i++)
        {
            var region = Regions[i % Regions.Length];
            var name = string.Create(CultureInfo.InvariantCulture, $"Server-{i:000000}");
            var department = string.Create(CultureInfo.InvariantCulture, $"Department-{i % 25}");
            var notes = string.Create(CultureInfo.InvariantCulture, $"Benchmark row {i}");

            switch (shape)
            {
                case CsvBenchmarkShape.Quoted:
                    name = string.Create(CultureInfo.InvariantCulture, $"Server,{i:000000}");
                    department = string.Create(CultureInfo.InvariantCulture, $"Department \"{i % 25}\"");
                    notes = string.Create(CultureInfo.InvariantCulture, $"Benchmark row {i}, \"quoted\", region {region}");
                    break;
                case CsvBenchmarkShape.Multiline:
                    notes = string.Create(CultureInfo.InvariantCulture, $"Benchmark row {i}\ncontinued value {i % 10}");
                    break;
            }

            rows[i - 1] = new CsvBenchmarkRow
            {
                Id = i,
                Name = name,
                Department = department,
                Region = region,
                IsEnabled = i % 3 != 0,
                Created = new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMinutes(i),
                Score = Math.Round((decimal)((i * 1.137) % 1000), 3),
                Owner = string.Create(CultureInfo.InvariantCulture, $"owner{i % 250}@example.test"),
                TicketCount = i % 17,
                Notes = notes
            };
        }

        return rows;
    }
}
