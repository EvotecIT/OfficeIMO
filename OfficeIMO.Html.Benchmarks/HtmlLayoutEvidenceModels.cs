namespace OfficeIMO.Html.Benchmarks;

internal sealed record HtmlLayoutEvidenceMeasurement(
    string Workload, int Iteration, int InputBytes, int PageCount, int TextCharacters,
    double ElapsedMilliseconds, long AllocatedBytes, long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes, long PeakWorkingSetGrowthBytes, long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record HtmlLayoutEvidenceSummary(
    string Workload, double MedianElapsedMilliseconds, double MedianAllocatedBytes,
    double MedianRetainedManagedHeapGrowthBytes, double MedianPeakManagedHeapGrowthBytes,
    double MedianAbsoluteProcessPeakWorkingSetBytes, int InputBytes, int PageCount, int TextCharacters);

internal sealed record HtmlLayoutEvidenceReport(
    DateTimeOffset MeasuredAtUtc, string SourceCommit, bool SourceTreeDirty, string Framework,
    string OperatingSystem, string Architecture, int ProcessorCount, int Repeat,
    IReadOnlyList<HtmlLayoutEvidenceMeasurement> Measurements, IReadOnlyList<HtmlLayoutEvidenceSummary> Summaries,
    IReadOnlyList<string> Failures);

internal sealed class HtmlLayoutBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<HtmlLayoutBudget> Budgets { get; set; } = new();
}

internal sealed class HtmlLayoutBudget {
    public string Workload { get; set; } = string.Empty;
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
    public long MaxRetainedManagedHeapGrowthBytes { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxAbsoluteProcessPeakWorkingSetBytes { get; set; }
}

internal readonly record struct HtmlLayoutMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);
