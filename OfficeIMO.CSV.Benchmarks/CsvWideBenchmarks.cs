#nullable enable

using System.Globalization;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using CsvHelperReader = CsvHelper.CsvReader;
using CsvHelperWriter = CsvHelper.CsvWriter;
using SylvanCsvDataReader = Sylvan.Data.Csv.CsvDataReader;

namespace OfficeIMO.CSV.Benchmarks;

[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net80)]
public class CsvWideBenchmarks
{
    private static readonly string[] Headers = CreateHeaders();

    private object?[][] _rows = [];
    private string _csvText = string.Empty;

    [Params(1000, 10000, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _rows = CsvWideBenchmarkData.Create(RowCount);

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
    public int OfficeIMO_ReadRowsReusableCallback()
    {
        using var reader = new StringReader(_csvText);
        var fieldCount = 0;
        CsvDocument.ReadRowsReusable(reader, (_, values) =>
        {
            fieldCount += values.Count;
        });

        return fieldCount;
    }

    [Benchmark]
    public int OfficeIMO_ReadRecordsReusableSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var fieldCount = 0;
        CsvDocument.ReadRecordsReusable(
            reader,
            values =>
            {
                fieldCount += values.Count;
            },
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return fieldCount;
    }

    [Benchmark]
    public int OfficeIMO_ReadFieldSpansSkipHeader()
    {
        using var reader = new StringReader(_csvText);
        var fieldCount = 0;
        CsvDocument.ReadFieldSpans(
            reader,
            (_, _, _) =>
            {
                fieldCount++;
            },
            new CsvLoadOptions { SkipInitialRecords = 1 });

        return fieldCount;
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

        return visitor.FieldCount;
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
                _ = csv.GetField(i);
                fieldCount++;
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
                _ = csv.GetString(i);
                fieldCount++;
            }
        }

        return fieldCount;
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
                _ = csv.GetFieldSpan(i);
                fieldCount++;
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

    private struct CountingFieldSpanVisitor : ICsvFieldSpanVisitor
    {
        public int FieldCount { get; private set; }

        public void VisitField(int recordIndex, int fieldIndex, ReadOnlySpan<char> value)
        {
            FieldCount++;
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
