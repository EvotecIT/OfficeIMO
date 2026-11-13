using System.Diagnostics;
using System.Text;

namespace OfficeIMO.Data.Conversions.Benchmarks;

/// <summary>A quick stopwatch run of single lanes (<c>--profile [name filter]</c>) for iterating on one path without a full BenchmarkDotNet job.</summary>
internal static class Profile {
    public static void Run(string? filter = null) {
        _filter = filter;
        byte[] nested = Encoding.UTF8.GetBytes(Workloads.NestedJson(20_000));
        byte[] flat = Encoding.UTF8.GetBytes(Workloads.FlatJson(20_000));
        Time("JSON->XML", () => ConversionBenchmarks.Run(nested, DataFormat.Json, DataFormat.Xml));
        Time("JSON->CSV", () => ConversionBenchmarks.Run(flat, DataFormat.Json, DataFormat.Csv));
        Time("JSON->JSON (parse+write floor)", () => ConversionBenchmarks.Run(flat, DataFormat.Json, DataFormat.Json));
        var lanes = new ConversionBenchmarks { Records = 20_000 };
        lanes.Setup();
        Time("JSON->XML Newtonsoft", lanes.JsonToXmlNewtonsoft);
        Time("JSON->CSV JsonDocument+CsvHelper", lanes.JsonToCsvCsvHelper);
        Time("CSV->JSON", lanes.CsvToJsonOfficeIMO);
        Time("CSV->JSON CsvHelper+STJ", lanes.CsvToJsonCsvHelper);
        Time("JSON->TSV", lanes.JsonToTsvOfficeIMO);
        Time("JSON->TSV JsonDocument+CsvHelper", lanes.JsonToTsvCsvHelper);
        Time("TSV->JSON", lanes.TsvToJsonOfficeIMO);
        Time("TSV->JSON CsvHelper+STJ", lanes.TsvToJsonCsvHelper);
        Time("TOML->JSON", lanes.TomlToJsonOfficeIMO);
        Time("TOML->JSON Tomlyn+Utf8JsonWriter", lanes.TomlToJsonTomlyn);
        Time("JSON->TOML", lanes.JsonToTomlOfficeIMO);
        Time("JSON->TOML JsonDocument+Tomlyn", lanes.JsonToTomlTomlyn);
    }

    private static string? _filter;

    private static void Time(string name, Func<long> run) {
        if (_filter != null && name.IndexOf(_filter, StringComparison.OrdinalIgnoreCase) < 0) return;
        for (int i = 0; i < 5; i++) run();
        // The minimum of many rounds is the most stable figure on a shared machine.
        const int rounds = 30;
        double ms = double.MaxValue;
        long before = GC.GetAllocatedBytesForCurrentThread();
        for (int i = 0; i < rounds; i++) {
            var timer = Stopwatch.StartNew();
            run();
            ms = Math.Min(ms, timer.Elapsed.TotalMilliseconds);
        }

        double mb = (GC.GetAllocatedBytesForCurrentThread() - before) / rounds / 1048576.0;
        Console.WriteLine($"{name,-32} {ms,8:F2} ms {mb,8:F2} MB");
    }
}
