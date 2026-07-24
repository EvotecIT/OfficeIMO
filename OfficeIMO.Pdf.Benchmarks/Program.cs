using System.Text.Json;

PdfPerformanceBudget budget = JsonSerializer.Deserialize<PdfPerformanceBudget>(
    File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "pdf-performance-budgets.json")),
    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
    ?? throw new InvalidOperationException("PDF performance budget manifest is invalid.");

byte[] corpus = PdfBenchmarkCorpus.Create();
IReadOnlyList<PdfPerformanceMeasurement> measurements = PdfBenchmarkRunner.Measure(corpus);
PdfPerformanceMeasurement cold = measurements.Single(measurement => measurement.Name == PdfBenchmarkRunner.AnalysisCold);
PdfPerformanceMeasurement cached = measurements.Single(measurement => measurement.Name == PdfBenchmarkRunner.AnalysisCached);
double speedup = cold.ElapsedMilliseconds / Math.Max(cached.ElapsedMilliseconds, 0.001D);
double allocationReduction = cold.AllocatedBytes / (double)Math.Max(cached.AllocatedBytes, 1L);

Console.WriteLine($"Corpus: {corpus.Length:N0} bytes, {PdfBenchmarkCorpus.PageCount} mixed pages");
foreach (PdfPerformanceMeasurement measurement in measurements) {
    Console.WriteLine(
        $"{measurement.Name,-18} {measurement.ElapsedMilliseconds,8:F1} ms " +
        $"{measurement.AllocatedBytes / 1048576D,8:F1} MiB allocated " +
        $"{measurement.PeakManagedHeapBytes / 1048576D,8:F1} MiB managed-peak " +
        $"{measurement.Output,12:N0} output");
    if (measurement.LargestSerializedObjectBytes > 0L) {
        Console.WriteLine(
            $"  writer={(measurement.IsForwardOnlyObjectSerialization ? "forward" : "buffered"),-8} " +
            $"page-peak={measurement.PeakRetainedPageContentBytes,12:N0} " +
            $"object-peak={measurement.PeakRetainedObjectBytes,12:N0} " +
            $"largest-object={measurement.LargestSerializedObjectBytes,12:N0} " +
            $"largest-transient={measurement.LargestTransientBufferBytes,12:N0}");
    }
}
Console.WriteLine($"Cached speedup: {speedup:F2}x");
Console.WriteLine($"Cached allocation reduction: {allocationReduction:F2}x");

if (!args.Contains("--verify-budgets", StringComparer.OrdinalIgnoreCase)) {
    return 0;
}

var failures = new List<string>();
foreach (PdfPerformanceMeasurement measurement in measurements) {
    if (!budget.Workloads.TryGetValue(measurement.Name, out PdfWorkloadBudget? workloadBudget)) {
        failures.Add(measurement.Name + ": no workload budget is defined.");
        continue;
    }

    if (measurement.ElapsedMilliseconds > workloadBudget.MaxElapsedMilliseconds) {
        failures.Add(
            $"{measurement.Name}: {measurement.ElapsedMilliseconds:F1} ms exceeded " +
            $"{workloadBudget.MaxElapsedMilliseconds:F1} ms.");
    }

    if (measurement.AllocatedBytes > workloadBudget.MaxAllocatedBytes) {
        failures.Add(
            $"{measurement.Name}: {measurement.AllocatedBytes:N0} allocated bytes exceeded " +
            $"{workloadBudget.MaxAllocatedBytes:N0}.");
    }

    if (measurement.PeakManagedHeapBytes > workloadBudget.MaxPeakManagedHeapBytes) {
        failures.Add(
            $"{measurement.Name}: {measurement.PeakManagedHeapBytes:N0} peak managed heap bytes exceeded " +
            $"{workloadBudget.MaxPeakManagedHeapBytes:N0}.");
    }
}

foreach (string workloadName in budget.Workloads.Keys) {
    if (!measurements.Any(measurement => measurement.Name == workloadName)) {
        failures.Add(workloadName + ": budget has no measured workload.");
    }
}

if (speedup < budget.MinimumCachedSpeedup) {
    failures.Add($"Cached workflow speedup {speedup:F2}x was below {budget.MinimumCachedSpeedup:F2}x.");
}

if (allocationReduction < budget.MinimumCachedAllocationReduction) {
    failures.Add(
        $"Cached workflow allocation reduction {allocationReduction:F2}x was below " +
        $"{budget.MinimumCachedAllocationReduction:F2}x.");
}

foreach (string failure in failures) {
    Console.Error.WriteLine("BUDGET FAILURE: " + failure);
}

return failures.Count == 0 ? 0 : 1;
