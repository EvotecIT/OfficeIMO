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
    IReadOnlyList<HtmlLayoutEvidenceMeasurement> Measurements, IReadOnlyList<HtmlLayoutEvidenceSummary> Summaries);

internal readonly record struct HtmlLayoutMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);
