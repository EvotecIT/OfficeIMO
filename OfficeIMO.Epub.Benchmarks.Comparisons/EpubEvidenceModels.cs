namespace OfficeIMO.Epub.Benchmarks.Comparisons;

internal sealed record EpubEvidenceMeasurement(
    string Engine,
    string Scale,
    int Iteration,
    int Operations,
    int RetainedDocuments,
    long InputBytes,
    int ChapterCount,
    long ContentCharacters,
    long TextCharacters,
    double ElapsedMilliseconds,
    double ElapsedMicrosecondsPerOperation,
    long AllocatedBytes,
    double AllocatedBytesPerOperation,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    string PathHash,
    string ContentHash,
    string TextHash);

internal sealed record EpubEvidenceSummary(
    string Scale,
    double ElapsedRatio,
    double AllocationRatio,
    double? RetainedManagedRatio,
    double? PeakManagedHeapRatio,
    double? ProcessPeakWorkingSetRatio,
    double OfficeElapsedMicrosecondsPerOperation,
    double VersOneElapsedMicrosecondsPerOperation,
    double OfficeAllocatedBytesPerOperation,
    double VersOneAllocatedBytesPerOperation);

internal sealed record EpubEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<string> ValidatedEquivalentScales,
    IReadOnlyList<EpubEvidenceMeasurement> Measurements,
    IReadOnlyList<EpubEvidenceSummary> Summaries);

internal readonly record struct EpubMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);
