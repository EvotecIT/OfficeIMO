using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.Benchmarks;

internal static class InvoicePerformanceBudgetRunner {
    internal static async Task<int> RunAsync(string[] args) {
        int iterations = IntegerOption(args, "--iterations", 5);
        int warmups = IntegerOption(args, "--warmups", 1);
        string? output = StringOption(args, "--output");
        if (iterations < 3) throw new ArgumentOutOfRangeException(nameof(args), "At least three measured iterations are required.");
        if (warmups < 0) throw new ArgumentOutOfRangeException(nameof(args), "Warmup count cannot be negative.");

        InvoiceBenchmarkCorpus corpus = InvoiceBenchmarkCorpus.Create();
        InvoiceValidator validator = InvoiceBenchmarkAuthority.CreateValidator();
        PdfInvoiceDocument pdfSnapshot = corpus.CapturePdf();
        var workloads = new[] {
            new Workload("xml-read", () => Task.FromResult<object>(InvoiceParser.Read(corpus.Xml)),
                value => InvoiceBenchmarkCorrectness.Read((InvoiceReadResult)value, corpus)),
            new Workload("xml-write", () => Task.FromResult<object>(InvoiceSerializer.Write(corpus.Invoice, InvoiceBenchmarkCorpus.Contract)),
                value => InvoiceBenchmarkCorrectness.Write((byte[])value, corpus)),
            new Workload("rules-validation", async () => await validator.ValidateAsync(corpus.Xml,
                    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2).ConfigureAwait(false),
                value => InvoiceBenchmarkCorrectness.Rules((InvoiceValidationReport)value)),
            new Workload("pdf-generation", () => Task.FromResult<object>(pdfSnapshot.ToPdfBytes(corpus.PdfOptions)),
                value => InvoiceBenchmarkCorrectness.Pdf((byte[])value, corpus.Xml))
        };

        foreach (Workload workload in workloads) {
            for (int index = 0; index < warmups; index++) workload.Validate(await workload.Execute().ConfigureAwait(false));
        }

        var results = new List<InvoicePerformanceResult>();
        foreach (Workload workload in workloads) {
            var elapsed = new double[iterations];
            var allocated = new long[iterations];
            for (int index = 0; index < iterations; index++) {
                long allocationStart = GC.GetTotalAllocatedBytes(precise: true);
                long start = Stopwatch.GetTimestamp();
                object result = await workload.Execute().ConfigureAwait(false);
                elapsed[index] = Stopwatch.GetElapsedTime(start).TotalMilliseconds;
                allocated[index] = GC.GetTotalAllocatedBytes(precise: true) - allocationStart;
                workload.Validate(result);
            }
            results.Add(new InvoicePerformanceResult(workload.Name, Median(elapsed), Median(allocated), elapsed, allocated));
        }

        string platform = OperatingSystem.IsWindows() ? "windows" : OperatingSystem.IsLinux() ? "linux" : OperatingSystem.IsMacOS() ? "macos" : "unknown";
        InvoicePerformanceBudgets budgets = JsonSerializer.Deserialize<InvoicePerformanceBudgets>(
            File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "invoice-performance-budgets.json")),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidDataException("Invoice performance budget manifest is invalid.");
        if (!budgets.Platforms.TryGetValue(platform, out Dictionary<string, InvoicePerformanceLimit>? platformBudgets))
            throw new PlatformNotSupportedException("No invoice performance budgets are defined for " + platform + ".");

        var failures = new List<string>();
        foreach (InvoicePerformanceResult result in results) {
            if (!platformBudgets.TryGetValue(result.Name, out InvoicePerformanceLimit? limit)) {
                failures.Add(result.Name + ": no " + platform + " budget is defined.");
                continue;
            }
            if (result.MedianElapsedMilliseconds > limit.MaxElapsedMilliseconds)
                failures.Add($"{result.Name}: {result.MedianElapsedMilliseconds:F3} ms exceeded {limit.MaxElapsedMilliseconds:F3} ms.");
            if (result.MedianAllocatedBytes > limit.MaxAllocatedBytes)
                failures.Add($"{result.Name}: {result.MedianAllocatedBytes:N0} allocated bytes exceeded {limit.MaxAllocatedBytes:N0}.");
        }
        foreach (string budgetName in platformBudgets.Keys)
            if (results.All(result => !string.Equals(result.Name, budgetName, StringComparison.Ordinal))) failures.Add(budgetName + ": budget has no measured workload.");

        var report = new InvoicePerformanceReport(
            platform,
            System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
            Environment.Version.ToString(),
            InvoiceBenchmarkCorpus.LineCount,
            corpus.Xml.Length,
            iterations,
            warmups,
            results,
            failures);
        string json = JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true });
        Console.WriteLine(json);
        if (!string.IsNullOrWhiteSpace(output)) {
            string fullPath = Path.GetFullPath(output!);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            File.WriteAllText(fullPath, json);
        }
        foreach (string failure in failures) Console.Error.WriteLine("BUDGET FAILURE: " + failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static int IntegerOption(string[] args, string name, int fallback) {
        string? value = StringOption(args, name);
        return value == null ? fallback : int.TryParse(value, out int parsed) ? parsed : throw new ArgumentException(name + " requires an integer.");
    }

    private static string? StringOption(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 == args.Length) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }

    private static double Median(double[] values) {
        Array.Sort(values);
        int middle = values.Length / 2;
        return (values.Length & 1) == 0 ? (values[middle - 1] + values[middle]) / 2D : values[middle];
    }

    private static long Median(long[] values) {
        Array.Sort(values);
        int middle = values.Length / 2;
        return (values.Length & 1) == 0 ? (values[middle - 1] + values[middle]) / 2L : values[middle];
    }

    private sealed record Workload(string Name, Func<Task<object>> Execute, Action<object> Validate);
}

internal sealed record InvoicePerformanceResult(string Name, double MedianElapsedMilliseconds, long MedianAllocatedBytes,
    IReadOnlyList<double> ElapsedMilliseconds, IReadOnlyList<long> AllocatedBytes);
internal sealed record InvoicePerformanceReport(string Platform, string Architecture, string Runtime, int InvoiceLines, int XmlBytes,
    int Iterations, int Warmups, IReadOnlyList<InvoicePerformanceResult> Workloads, IReadOnlyList<string> Failures);
internal sealed class InvoicePerformanceBudgets {
    public Dictionary<string, Dictionary<string, InvoicePerformanceLimit>> Platforms { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}
internal sealed class InvoicePerformanceLimit {
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
}
