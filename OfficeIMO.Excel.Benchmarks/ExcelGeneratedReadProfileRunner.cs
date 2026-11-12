#if NET10_0_OR_GREATER
using System.Globalization;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Repeats the validated generated XLSX operations for an external CPU profiler.</summary>
internal static class ExcelGeneratedReadProfileRunner {
    /// <summary>Creates validated, fixed inputs before comparing different reader builds.</summary>
    internal static void PrepareInputs(string directory) {
        if (!string.IsNullOrEmpty(Environment.GetEnvironmentVariable("OFFICEIMO_EXCEL_GENERATED_INPUT"))) {
            throw new InvalidOperationException("Unset OFFICEIMO_EXCEL_GENERATED_INPUT before preparing inputs.");
        }
        Directory.CreateDirectory(directory);
        foreach (ExcelFileFormat format in new[] { ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls }) {
            foreach (bool escaped in new[] { false, true }) {
                var fixture = new ExcelReaderGeneratedBenchmarks { Format = format, RowCount = 50_000, EscapedNames = escaped };
                try {
                    fixture.Setup();
                    string name = $"generated-four-columns-{(escaped ? "escaped-" : string.Empty)}50000.{format.ToString().ToLowerInvariant()}";
                    File.Copy(fixture.InputPath, Path.Combine(directory, name), overwrite: false);
                } finally {
                    fixture.Cleanup();
                }
            }
        }
    }

    /// <summary>Arguments: read|parse|prepare, iteration count, and optionally escaped.</summary>
    internal static void Run(string[] arguments) {
        string operation = arguments.Length > 1 ? arguments[1] : "read";
        int iterations = arguments.Length > 2 ? int.Parse(arguments[2], CultureInfo.InvariantCulture) : 200;
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(iterations);
        var fixture = new ExcelReaderGeneratedBenchmarks {
            Format = arguments.Length > 4 ? Enum.Parse<ExcelFileFormat>(arguments[4], ignoreCase: true) : ExcelFileFormat.Xlsx,
            RowCount = 50_000,
            EscapedNames = arguments.Length > 3 && arguments[3] == "escaped"
        };
        try {
            fixture.Setup();
            Func<long> run = operation switch {
                "read" => fixture.OfficeIMO_Read,
                "parse" => fixture.OfficeIMO_Parse,
                "erread" => fixture.ExcelReaderNet_Read,
                "erparse" => fixture.ExcelReaderNet_Parse,
                "prepare" => () => {
                    using var reader = ExcelDocument.OpenReader(fixture.InputPath);
                    if (!reader.HasRows) throw new InvalidDataException("Missing rows.");
                    return reader.FieldCount;
                },
                _ => throw new ArgumentException("Expected read, parse or prepare.", nameof(arguments))
            };
            for (int i = 0; i < 16; i++) run();
            long observation = 0;
            var samples = new double[iterations];
            for (int i = 0; i < iterations; i++) {
                long start = System.Diagnostics.Stopwatch.GetTimestamp();
                observation ^= run();
                samples[i] = System.Diagnostics.Stopwatch.GetElapsedTime(start).TotalMilliseconds;
            }
            Array.Sort(samples);
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"Min {samples[0]:F3} P10 {samples[iterations / 10]:F3} Median {samples[iterations / 2]:F3} ms"));
            Console.WriteLine($"Completed {iterations} {operation} operations; observation {observation}.");
        } finally {
            fixture.Cleanup();
        }
    }
}
#endif
