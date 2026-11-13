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

    /// <summary>
    /// BenchmarkDotNet runs every format in one process, so shared reader code is tiered with profile data
    /// from the formats measured first. OFFICEIMO_PROFILE_PRETRAIN=xlsx,xlsb reproduces that order;
    /// OFFICEIMO_PROFILE_PRETRAIN_OPS limits it to some of parse, erparse, read and erread.
    /// </summary>
    private static void PretrainOtherFormats(ExcelFileFormat measured) {
        string? formats = Environment.GetEnvironmentVariable("OFFICEIMO_PROFILE_PRETRAIN");
        if (string.IsNullOrEmpty(formats)) return;
        foreach (string name in formats.Split(',', StringSplitOptions.RemoveEmptyEntries)) {
            var format = Enum.Parse<ExcelFileFormat>(name, ignoreCase: true);
            if (format == measured) continue;
            foreach (bool escaped in new[] { false, true }) {
                var other = new ExcelReaderGeneratedBenchmarks { Format = format, RowCount = 50_000, EscapedNames = escaped };
                try {
                    other.Setup();
                    string[] operations = (Environment.GetEnvironmentVariable("OFFICEIMO_PROFILE_PRETRAIN_OPS")
                        ?? "parse,erparse,read,erread").Split(',');
                    for (int i = 0; i < 150; i++) {
                        if (operations.Contains("parse")) other.OfficeIMO_Parse();
                        if (operations.Contains("erparse")) other.ExcelReaderNet_Parse();
                        if (operations.Contains("read")) other.OfficeIMO_Read();
                        if (operations.Contains("erread")) other.ExcelReaderNet_Read();
                    }
                } finally {
                    other.Cleanup();
                }
            }
        }
    }

    // Cycles charged to the calling thread exclude time descheduled by other load on the host; work on
    // other threads is not included, so compare multithreaded paths by wall time.
    private static ulong ThreadCycles() => OperatingSystem.IsWindows() && QueryThreadCycleTime(GetCurrentThread(), out ulong value) ? value : 0;

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern bool QueryThreadCycleTime(IntPtr thread, out ulong cycles);

    [System.Runtime.InteropServices.DllImport("kernel32.dll")]
    private static extern IntPtr GetCurrentThread();

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
            PretrainOtherFormats(fixture.Format);
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
            int gen0 = GC.CollectionCount(0), gen1 = GC.CollectionCount(1), gen2 = GC.CollectionCount(2);
            var samples = new double[iterations];
            var cycles = new double[iterations];
            for (int i = 0; i < iterations; i++) {
                ulong startCycles = ThreadCycles();
                long start = System.Diagnostics.Stopwatch.GetTimestamp();
                observation ^= run();
                samples[i] = System.Diagnostics.Stopwatch.GetElapsedTime(start).TotalMilliseconds;
                cycles[i] = (ThreadCycles() - startCycles) / 1e6;
            }
            Array.Sort(cycles);
            double mean = samples.Average();
            Array.Sort(samples);
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"Min {samples[0]:F3} P10 {samples[iterations / 10]:F3} Median {samples[iterations / 2]:F3} Mean {mean:F3} P90 {samples[iterations * 9 / 10]:F3} ms; calling-thread Mcycles Min {cycles[0]:F2} P10 {cycles[iterations / 10]:F2} Median {cycles[iterations / 2]:F2}; GC {GC.CollectionCount(0) - gen0}/{GC.CollectionCount(1) - gen1}/{GC.CollectionCount(2) - gen2}"));
            Console.WriteLine($"Completed {iterations} {operation} operations; observation {observation}.");
        } finally {
            fixture.Cleanup();
        }
    }
}
#endif
