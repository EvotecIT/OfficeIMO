namespace OfficeIMO.Email.Benchmarks.Comparisons;

internal sealed record EmailMimeEvidenceMeasurement(
    string Operation,
    string Engine,
    string Scale,
    int Iteration,
    int Operations,
    int RetainedResults,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytes,
    double AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    string SemanticFingerprint);

internal sealed record EmailMimeEvidenceSummary(
    string Operation,
    string Scale,
    double ElapsedRatio,
    double AllocationRatio,
    double? RetainedManagedRatio,
    double? PeakManagedHeapRatio,
    double? ProcessPeakWorkingSetRatio,
    double? OutputSizeRatio);

internal sealed record EmailMimeEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<string> ValidatedEquivalentScales,
    IReadOnlyList<EmailMimeEvidenceMeasurement> Measurements,
    IReadOnlyList<EmailMimeEvidenceSummary> Summaries);

internal readonly record struct EmailMimeMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);
