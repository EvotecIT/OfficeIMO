using System.Globalization;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>Generated CSV shapes shared by the opt-in comparison runners.</summary>
public enum CsvBenchmarkShape
{
    Mixed,
    Quoted,
    Multiline
}

/// <summary>Mixed values used by the generated CSV read and write comparisons.</summary>
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

/// <summary>Owns the generated mixed CSV records and their independent expected values.</summary>
internal static class CsvBenchmarkData
{
    internal static readonly string[] Headers =
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

    /// <summary>Creates the same LF-separated input for both generated comparison runners.</summary>
    internal static string CreateText(IEnumerable<CsvBenchmarkRow> rows)
    {
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteObjects(writer, rows, new CsvSaveOptions { NewLine = "\n" });
        return writer.ToString();
    }

    /// <summary>Projects values in the shared header order.</summary>
    internal static object?[] ProjectRow(CsvBenchmarkRow row)
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

    /// <summary>Builds expected decoded values with invariant CLR formatting.</summary>
    internal static string?[] ProjectTextRow(object?[] row)
    {
        var values = new string?[row.Length];
        for (var i = 0; i < row.Length; i++)
        {
            values[i] = Convert.ToString(row[i], CultureInfo.InvariantCulture);
        }

        return values;
    }

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
