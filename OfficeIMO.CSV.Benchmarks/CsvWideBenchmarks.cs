#nullable enable

using System.Globalization;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using nietras.SeparatedValues;
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
    private static readonly SepReaderOptions SepReadOptions = SepLib.New(',').Reader(options => options with { Unescape = true });
    private static readonly SepWriterOptions SepWriteOptions = SepLib.New(',').Writer(options => options with { WriteHeader = true, Escape = true });
    private static readonly SylvanCsvDataWriterOptions SylvanWriterOptions = new() { NewLine = "\n" };

    private object?[][] _rows = [];
    private string?[][] _textRows = [];
    private string _csvText = string.Empty;

    [Params(1000, 10000, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _rows = CsvWideBenchmarkData.Create(RowCount);
        _textRows = _rows.Select(ProjectTextRow).ToArray();

        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        foreach (var row in _rows)
        {
            csv.WriteRow(Headers, row);
        }

        _csvText = writer.ToString();
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        foreach (var row in _rows)
        {
            csv.WriteRow(Headers, row);
        }

        return writer.GetStringBuilder().Length;
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

        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int CsvHelper_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CultureInfo.InvariantCulture);
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

        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int Sylvan_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows);
        using var csv = SylvanCsvDataWriter.Create(writer, SylvanWriterOptions);
        csv.Write(reader);
        return writer.GetStringBuilder().Length;
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

        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int Dataplat_WriteFromReader()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var reader = new BenchmarkArrayDataReader(Headers, _rows);
        using var csv = new DataplatCsvWriter(writer, DataplatWriterOptions);
        csv.WriteFromReader(reader);
        return writer.GetStringBuilder().Length;
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

        return csv.ToString().Length;
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
