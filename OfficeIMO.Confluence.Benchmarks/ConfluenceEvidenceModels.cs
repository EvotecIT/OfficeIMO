namespace OfficeIMO.Confluence.Benchmarks;

internal sealed record ConfluenceEvidenceMeasurement(
    int PageCharacters,
    int Iteration,
    long InputBytes,
    long OutputBytes,
    int OutputCharacters,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    string OriginalSha256,
    string UpdatedSha256);

internal sealed record ConfluenceEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<ConfluenceEvidenceMeasurement> Measurements);

internal readonly record struct ConfluenceMemoryPeak(long ManagedHeapBytes, long WorkingSetBytes);
