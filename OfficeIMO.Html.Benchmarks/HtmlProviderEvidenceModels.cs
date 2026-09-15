namespace OfficeIMO.Html.Benchmarks;

internal sealed record HtmlProviderEvidenceMeasurement(
    string Scenario,
    int Iteration,
    int InputCharacters,
    int ResultItems,
    string ResultFingerprintSha256,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long RetainedManagedHeapBytesPerResult,
    long AbsoluteProcessPeakWorkingSetBytes);

internal sealed record HtmlProviderEvidenceSummary(
    string Scenario,
    double MedianElapsedMilliseconds,
    double MedianAllocatedBytes,
    double MedianRetainedManagedHeapBytesPerResult,
    double MedianAbsoluteProcessPeakWorkingSetBytes,
    int InputCharacters,
    int ResultItems);

internal sealed record HtmlProviderEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool TrackedSourceDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<HtmlProviderAssemblyEvidence> Providers,
    IReadOnlyList<HtmlProviderEvidenceMeasurement> Measurements,
    IReadOnlyList<HtmlProviderEvidenceSummary> Summaries);

internal sealed record HtmlProviderAssemblyEvidence(
    string Assembly,
    string Version,
    string InformationalVersion,
    long FileBytes);
