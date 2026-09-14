namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record HtmlStaticBudgetReport(
    int SchemaVersion,
    DateTimeOffset GeneratedUtc,
    HtmlStaticBudgetEnvironment Environment,
    HtmlStaticBudgetSource Source,
    HtmlStaticBudgetCeiling? Ceiling,
    HtmlStaticBudgetRun Cold,
    HtmlStaticBudgetRun Warm,
    HtmlStaticBudgetCancellation Cancellation,
    IReadOnlyList<string> Failures);

internal sealed record HtmlStaticBudgetEnvironment(
    string OsFamily,
    string OsDescription,
    string OsArchitecture,
    string ProcessArchitecture,
    string FrameworkDescription,
    string MachineName,
    int ProcessorCount);

internal sealed record HtmlStaticBudgetSource(
    string? Commit,
    bool? WorktreeClean,
    string CorpusId,
    string ManifestSha256,
    int CaseCount,
    int OutputsPerIteration);

internal sealed record HtmlStaticBudgetRun(
    string Mode,
    double ProcessElapsedMilliseconds,
    long PeakWorkingSetBytes,
    int MemorySampleCount,
    IReadOnlyList<HtmlStaticBudgetIteration> Iterations,
    bool Deterministic);

internal sealed record HtmlStaticBudgetIteration(
    int Number,
    double ElapsedMilliseconds,
    long ManagedAllocatedBytes,
    long OutputBytes,
    int CaseCount,
    int OutputCount,
    string FingerprintSha256);

internal sealed record HtmlStaticBudgetCancellation(
    string Status,
    double ElapsedMilliseconds,
    string Detail);

internal sealed record HtmlPdfProcessTreeMemoryEvidence(
    long PeakWorkingSetBytes,
    int SampleCount,
    int MinimumObservedProcessCount,
    int MaximumObservedProcessCount,
    string Sampler);

internal sealed record HtmlStaticBudgetCeiling(
    string OsFamily,
    double ColdProcessElapsedMilliseconds,
    double WarmIterationElapsedMilliseconds,
    long ColdManagedAllocatedBytes,
    long WarmManagedAllocatedBytes,
    long PeakWorkingSetBytes,
    long OutputBytesPerIteration,
    double CancellationElapsedMilliseconds);

internal sealed class HtmlStaticBudgetConfiguration {
    public int SchemaVersion { get; set; }
    public string CorpusId { get; set; } = string.Empty;
    public string ManifestSha256 { get; set; } = string.Empty;
    public int WarmIterations { get; set; }
    public List<HtmlStaticBudgetCeiling> Platforms { get; set; } = new();
}
