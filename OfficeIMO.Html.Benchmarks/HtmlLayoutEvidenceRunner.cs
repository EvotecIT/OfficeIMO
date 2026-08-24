using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Benchmarks;

internal static class HtmlLayoutEvidenceRunner {
    private static readonly string[] Workloads = [
        "Report100", "Purchase250", "Purchase2500", "Long100", "Long1000", "StaticStandards"
    ];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 1) {
            Console.Error.WriteLine("Usage: --layout-evidence-probe <workload>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0]), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        try {
            string? workloadFilter = GetOption(args, "--workload");
            string? jsonPath = GetOption(args, "--json");
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            string[] workloads = string.IsNullOrWhiteSpace(workloadFilter)
                ? Workloads
                : [ResolveWorkload(workloadFilter!)];
            var measurements = new List<HtmlLayoutEvidenceMeasurement>(workloads.Length * repeat);
            foreach (string workload in workloads) {
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    HtmlLayoutEvidenceMeasurement measurement = RunChildProbe(workload) with { Iteration = iteration };
                    measurements.Add(measurement);
                    Console.WriteLine(
                        $"{workload,-16} #{iteration} {measurement.ElapsedMilliseconds,10:F2} ms " +
                        $"{measurement.AllocatedBytes / 1048576D,9:F2} MiB alloc " +
                        $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,9:F2} MiB retained " +
                        $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,9:F2} MiB managed peak " +
                        $"{measurement.AbsoluteProcessPeakWorkingSetBytes / 1048576D,9:F2} MiB process peak " +
                        $"{measurement.PageCount,5:N0} pages");
                }
            }
            IReadOnlyList<HtmlLayoutEvidenceSummary> summaries = workloads.Select(workload => {
                HtmlLayoutEvidenceMeasurement[] values = measurements.Where(value => value.Workload == workload).ToArray();
                return new HtmlLayoutEvidenceSummary(
                    workload,
                    Median(values.Select(value => value.ElapsedMilliseconds)),
                    Median(values.Select(value => (double)value.AllocatedBytes)),
                    Median(values.Select(value => (double)value.RetainedManagedHeapGrowthBytes)),
                    Median(values.Select(value => (double)value.PeakManagedHeapGrowthBytes)),
                    Median(values.Select(value => (double)value.AbsoluteProcessPeakWorkingSetBytes)),
                    values[0].InputBytes,
                    values[0].PageCount,
                    values[0].TextCharacters);
            }).ToArray();
            var report = new HtmlLayoutEvidenceReport(
                DateTimeOffset.UtcNow, ResolveCommit(), ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription, RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(), Environment.ProcessorCount,
                repeat, measurements, summaries);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static HtmlLayoutEvidenceMeasurement Measure(string workloadName) {
        string workload = ResolveWorkload(workloadName);
        HtmlLayoutScenario scenario = HtmlLayoutScenario.Create(workload);
        int warmups = workload is "Long1000" or "Purchase2500" ? 1 : 3;
        for (int index = 0; index < warmups; index++) GC.KeepAlive(scenario.Render());

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        HtmlRenderDocument timed = scenario.Render();
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        scenario.Validate(timed);
        GC.KeepAlive(timed);

        timed = null!;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        using var sampler = new HtmlLayoutMemorySampler(process);
        HtmlRenderDocument retained = scenario.Render();
        scenario.Validate(retained);
        HtmlLayoutMemoryPeak peak = sampler.Stop();
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retainedGrowth = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long processPeak = process.PeakWorkingSet64;
        GC.KeepAlive(retained);

        return new HtmlLayoutEvidenceMeasurement(
            workload, 1, scenario.InputBytes, retained.Pages.Count, retained.Text.Length,
            stopwatch.Elapsed.TotalMilliseconds, allocated, retainedGrowth,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore), processPeak);
    }

    private static HtmlLayoutEvidenceMeasurement RunChildProbe(string workload) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve process path.");
        var info = new ProcessStartInfo { FileName = processPath, RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            info.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        info.ArgumentList.Add("--layout-evidence-probe");
        info.ArgumentList.Add(workload);
        using Process child = Process.Start(info) ?? throw new InvalidOperationException("Unable to start HTML layout probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {workload} failed: {error}");
        return JsonSerializer.Deserialize<HtmlLayoutEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {workload} returned no measurement.");
    }

    private static string ResolveWorkload(string value) => Workloads.FirstOrDefault(
        workload => string.Equals(workload, value, StringComparison.OrdinalIgnoreCase))
        ?? throw new ArgumentException("Unknown HTML layout workload: " + value);

    private static double Median(IEnumerable<double> values) {
        double[] ordered = values.OrderBy(value => value).ToArray();
        int middle = ordered.Length / 2;
        return ordered.Length % 2 == 0 ? (ordered[middle - 1] + ordered[middle]) / 2D : ordered[middle];
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length) throw new ArgumentException(name + " requires a value.");
        return args[index + 1];
    }

    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) {
        string? value = GetOption(args, name);
        return value == null ? defaultValue : int.TryParse(value, out int parsed) && parsed > 0
            ? parsed : throw new ArgumentException(name + " must be positive.");
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try { using Process process = Process.Start(GitInfo("rev-parse", "HEAD"))!; string output = process.StandardOutput.ReadToEnd().Trim(); process.WaitForExit(); return process.ExitCode == 0 ? output : "unknown"; }
        catch { return "unknown"; }
    }

    private static bool ResolveSourceTreeDirty() {
        try {
            using Process tracked = Process.Start(GitInfo("diff", "--quiet", "HEAD", "--"))!; tracked.WaitForExit(); if (tracked.ExitCode != 0) return true;
            using Process untracked = Process.Start(GitInfo("ls-files", "--others", "--exclude-standard"))!; string output = untracked.StandardOutput.ReadToEnd(); untracked.WaitForExit(); return untracked.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch { return true; }
    }

    private static ProcessStartInfo GitInfo(params string[] arguments) {
        var info = new ProcessStartInfo("git") { RedirectStandardOutput = true, RedirectStandardError = true, UseShellExecute = false, CreateNoWindow = true };
        foreach (string argument in arguments) info.ArgumentList.Add(argument);
        return info;
    }
}

internal sealed class HtmlLayoutScenario {
    private readonly string _html;
    private readonly HtmlConversionDocument? _prepared;
    private readonly HtmlRenderOptions _options;

    private HtmlLayoutScenario(string name, string html, HtmlRenderOptions options, bool endToEnd) {
        Name = name;
        _html = html;
        _options = options;
        _prepared = endToEnd ? null : HtmlConversionDocument.Parse(html);
        InputBytes = Encoding.UTF8.GetByteCount(html);
    }

    internal string Name { get; }
    internal int InputBytes { get; }
    internal HtmlRenderDocument Render() => HtmlRenderEngine.Render(_prepared ?? HtmlConversionDocument.Parse(_html), _options);

    internal void Validate(HtmlRenderDocument rendered) {
        if (rendered.Pages.Count == 0 || rendered.Text.Length == 0) throw new InvalidOperationException(Name + " produced no layout output.");
        string[] markers = Name switch {
            "Report100" => ["Benchmark Report", "Line 99"],
            "Purchase250" => ["SKU-00000", "SKU-00249", "Total $"],
            "Purchase2500" => ["SKU-00000", "SKU-02499", "Total $"],
            "Long100" => ["PAGE-0000", "PAGE-0099"],
            "Long1000" => ["PAGE-0000", "PAGE-0999"],
            _ => ["Static standards packet", "Second-page evidence"]
        };
        if (markers.Any(marker => !rendered.Text.Contains(marker, StringComparison.Ordinal))) {
            throw new InvalidOperationException(Name + " lost a required text marker.");
        }
        int? expectedPages = Name switch { "Long100" => 100, "Long1000" => 1000, "StaticStandards" => 2, _ => null };
        if (expectedPages.HasValue && rendered.Pages.Count != expectedPages.Value) {
            throw new InvalidOperationException($"{Name} produced {rendered.Pages.Count} pages; expected {expectedPages.Value}.");
        }
        if (Name == "StaticStandards") rendered.RequireNoLoss();
    }

    internal static HtmlLayoutScenario Create(string name) => name switch {
        "Report100" => new(name, HtmlBenchmarkCorpus.BuildReport(100), HtmlBenchmarkCorpus.CreateContinuousOptions(), true),
        "Purchase250" => new(name, HtmlBenchmarkCorpus.BuildPurchaseTable(250), HtmlBenchmarkCorpus.CreatePagedOptions(), false),
        "Purchase2500" => new(name, HtmlBenchmarkCorpus.BuildPurchaseTable(2500), HtmlBenchmarkCorpus.CreatePagedOptions(), false),
        "Long100" => new(name, HtmlBenchmarkCorpus.BuildLongDocument(100), LongOptions(100), false),
        "Long1000" => new(name, HtmlBenchmarkCorpus.BuildLongDocument(1000), LongOptions(1000), false),
        "StaticStandards" => new(name, HtmlBenchmarkCorpus.BuildStaticStandardsShowcase(), HtmlBenchmarkCorpus.CreateStaticStandardsOptions(), false),
        _ => throw new ArgumentException("Unknown HTML layout workload: " + name)
    };

    private static HtmlRenderOptions LongOptions(int pages) => new() {
        Mode = HtmlRenderMode.Paged, PageSize = OfficePageSizes.Letter,
        Margins = HtmlRenderMargins.All(48D), MaxPageCount = pages
    };
}
