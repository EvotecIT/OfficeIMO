namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record HtmlPdfProcessTreeMemoryEvidence(
    long PeakWorkingSetBytes,
    int SampleCount,
    int MinimumObservedProcessCount,
    int MaximumObservedProcessCount,
    string Sampler);

internal sealed record H10BudgetOperation(
    string Name,
    string Kind,
    bool Passed,
    bool HasLoss,
    double ConversionMilliseconds,
    long ManagedAllocatedBytes,
    double ProcessMilliseconds,
    HtmlPdfProcessTreeMemoryEvidence ProcessTreeMemory,
    long OutputBytes,
    string? OutputSha256,
    int? PageCount,
    string ReportPath);

internal sealed record H10BudgetLimit(
    string Name,
    double ConversionMilliseconds,
    long ManagedAllocatedBytes,
    long PeakWorkingSetBytes);

internal sealed record H10BudgetPlatform(
    string OsFamily,
    string CaseId,
    IReadOnlyList<H10BudgetLimit> Operations);

internal sealed record H10BudgetConfiguration(
    int SchemaVersion,
    string SourceSha256,
    string CaseSha256,
    IReadOnlyList<H10BudgetPlatform> Platforms);
