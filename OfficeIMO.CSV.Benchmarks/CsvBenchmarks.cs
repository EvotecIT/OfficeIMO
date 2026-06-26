#nullable enable

using System.Globalization;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using CsvHelperReader = CsvHelper.CsvReader;
using CsvHelperWriter = CsvHelper.CsvWriter;

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
    private string _csvText = string.Empty;

    [Params(1000, 10000, 25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup()
    {
        _rows = CsvBenchmarkData.Create(RowCount);
        _projectedRows = _rows.Select(ProjectRow).ToArray();

        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, _rows, new CsvSaveOptions { NewLine = "\n" });
        _csvText = writer.ToString();
    }

    [Benchmark(Baseline = true)]
    public int OfficeIMO_WriteObjects()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, _rows, new CsvSaveOptions { NewLine = "\n" });
        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int OfficeIMO_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvObjectWriter(writer, new CsvSaveOptions { NewLine = "\n" }, leaveOpen: true);
        foreach (object?[] row in _projectedRows)
        {
            csv.WriteRow(Headers, row);
        }

        return writer.GetStringBuilder().Length;
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

        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int CsvHelper_WriteTypedRecords()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CultureInfo.InvariantCulture);
        csv.WriteRecords(_rows);
        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int CsvHelper_WriteProjectedRows()
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        using var csv = new CsvHelperWriter(writer, CultureInfo.InvariantCulture);
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

        return writer.GetStringBuilder().Length;
    }

    [Benchmark]
    public int OfficeIMO_ReadRowsCallback()
    {
        using var reader = new StringReader(_csvText);
        var fieldCount = 0;
        CsvDocument.ReadRows(reader, (_, values) =>
        {
            fieldCount += values.Count;
        });

        return fieldCount;
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
    public int OfficeIMO_ReadStreamingRows()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.Stream });
        var fieldCount = 0;
        foreach (CsvRow row in document.AsEnumerable())
        {
            fieldCount += row.FieldCount;
        }

        return fieldCount;
    }

    [Benchmark]
    public int OfficeIMO_ReadInMemoryRows()
    {
        var document = CsvDocument.Parse(_csvText, new CsvLoadOptions { Mode = CsvLoadMode.InMemory });
        var fieldCount = 0;
        foreach (CsvRow row in document.AsEnumerable())
        {
            fieldCount += row.FieldCount;
        }

        return fieldCount;
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

    public static CsvBenchmarkRow[] Create(int count)
    {
        var rows = new CsvBenchmarkRow[count];
        for (var i = 1; i <= count; i++)
        {
            rows[i - 1] = new CsvBenchmarkRow
            {
                Id = i,
                Name = string.Create(CultureInfo.InvariantCulture, $"Server-{i:000000}"),
                Department = string.Create(CultureInfo.InvariantCulture, $"Department-{i % 25}"),
                Region = Regions[i % Regions.Length],
                IsEnabled = i % 3 != 0,
                Created = new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddMinutes(i),
                Score = Math.Round((decimal)((i * 1.137) % 1000), 3),
                Owner = string.Create(CultureInfo.InvariantCulture, $"owner{i % 250}@example.test"),
                TicketCount = i % 17,
                Notes = string.Create(CultureInfo.InvariantCulture, $"Benchmark row {i}")
            };
        }

        return rows;
    }
}
