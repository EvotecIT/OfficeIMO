using System.Diagnostics;
using System.Text.Json;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal static class PowerPointBenchmarkEvidence {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };

    internal static PowerPointBenchmarkBudgetManifest LoadBudgetManifest() {
        string path = Path.Combine(AppContext.BaseDirectory,
            "powerpoint-performance-budgets.json");
        return JsonSerializer.Deserialize<PowerPointBenchmarkBudgetManifest>(
                   File.ReadAllText(path), JsonOptions)
               ?? throw new InvalidOperationException(
                   "PowerPoint performance budget manifest is invalid.");
    }

    internal static void EvaluateBudget(
        PowerPointBenchmarkBudgetManifest manifest,
        PowerPointBaselineMeasurement measurement,
        ICollection<string> failures) {
        PowerPointBenchmarkBudget? budget = manifest.Budgets.FirstOrDefault(
            item => string.Equals(item.Operation, measurement.Operation,
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(item.Scale, measurement.Scale,
                        StringComparison.OrdinalIgnoreCase));
        if (budget == null) {
            failures.Add(
                $"Missing budget for {measurement.Operation}/{measurement.Scale}.");
            return;
        }

        string lane = measurement.Operation + "/" + measurement.Scale;
        if (measurement.ElapsedMilliseconds > budget.MaxElapsedMilliseconds) {
            failures.Add(
                $"{lane}: elapsed {measurement.ElapsedMilliseconds:F2} ms > {budget.MaxElapsedMilliseconds:F2} ms.");
        }
        if (measurement.AllocatedBytes > budget.MaxAllocatedBytes) {
            failures.Add(
                $"{lane}: allocations {measurement.AllocatedBytes} > {budget.MaxAllocatedBytes} bytes.");
        }
        if (measurement.PeakManagedHeapGrowthBytes
            > budget.MaxPeakManagedHeapGrowthBytes) {
            failures.Add(
                $"{lane}: managed-heap peak growth {measurement.PeakManagedHeapGrowthBytes} > {budget.MaxPeakManagedHeapGrowthBytes} bytes.");
        }
        if (measurement.PeakWorkingSetBytes
            > budget.MaxPeakWorkingSetBytes) {
            failures.Add(
                $"{lane}: process peak {measurement.PeakWorkingSetBytes} > {budget.MaxPeakWorkingSetBytes} bytes.");
        }
        if (measurement.OutputBytes > budget.MaxOutputBytes) {
            failures.Add(
                $"{lane}: output {measurement.OutputBytes} > {budget.MaxOutputBytes} bytes.");
        }
    }

    internal static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            using Process process = Process.Start(CreateGitStartInfo(
                "rev-parse", "HEAD"))!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch {
            return "unknown";
        }
    }

    internal static bool ResolveSourceTreeDirty() {
        try {
            using Process tracked = Process.Start(CreateGitStartInfo(
                "diff", "--quiet", "HEAD", "--"))!;
            tracked.WaitForExit();
            if (tracked.ExitCode != 0) return true;
            using Process untracked = Process.Start(CreateGitStartInfo(
                "ls-files", "--others", "--exclude-standard"))!;
            string output = untracked.StandardOutput.ReadToEnd();
            untracked.WaitForExit();
            return untracked.ExitCode != 0
                   || !string.IsNullOrWhiteSpace(output);
        } catch {
            return true;
        }
    }

    private static ProcessStartInfo CreateGitStartInfo(
        params string[] arguments) {
        var startInfo = new ProcessStartInfo {
            FileName = "git",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) {
            startInfo.ArgumentList.Add(argument);
        }
        return startInfo;
    }
}

internal sealed class PowerPointBenchmarkBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<PowerPointBenchmarkBudget> Budgets { get; set; } = new();
}

internal sealed class PowerPointBenchmarkBudget {
    public string Operation { get; set; } = string.Empty;
    public string Scale { get; set; } = string.Empty;
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxPeakWorkingSetBytes { get; set; }
    public long MaxOutputBytes { get; set; }
}

internal sealed class PowerPointManagedHeapSampler : IDisposable {
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _stop = new(false);
    private long _peakBytes;
    private int _stopped;

    internal PowerPointManagedHeapSampler() {
        _peakBytes = GC.GetTotalMemory(forceFullCollection: false);
        _thread = new Thread(SampleUntilStopped) {
            IsBackground = true,
            Name = "OfficeIMO.PowerPoint managed heap sampler"
        };
        _thread.Start();
    }

    internal long Stop() {
        if (Interlocked.Exchange(ref _stopped, 1) == 0) {
            _stop.Set();
            _thread.Join();
            RecordCurrentHeap();
        }
        return Interlocked.Read(ref _peakBytes);
    }

    public void Dispose() {
        Stop();
        _stop.Dispose();
    }

    private void SampleUntilStopped() {
        while (!_stop.Wait(1)) RecordCurrentHeap();
    }

    private void RecordCurrentHeap() {
        long observed = GC.GetTotalMemory(forceFullCollection: false);
        long current = Interlocked.Read(ref _peakBytes);
        while (observed > current) {
            long prior = Interlocked.CompareExchange(ref _peakBytes,
                observed, current);
            if (prior == current) return;
            current = prior;
        }
    }
}
