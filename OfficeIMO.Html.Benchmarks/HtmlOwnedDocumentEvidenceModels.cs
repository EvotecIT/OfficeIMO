namespace OfficeIMO.Html.Benchmarks;

internal sealed record HtmlOwnedDocumentEvidenceMeasurement(
    string Operation,
    string Scale,
    int Iteration,
    int InputCharacters,
    int ResultItems,
    string ResultFingerprintSha256,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long RetainedManagedHeapGrowthBytes,
    long PeakManagedHeapGrowthBytes,
    long AbsoluteProcessPeakWorkingSetBytes,
    int OutputCharacters);

internal sealed record HtmlOwnedDocumentEvidenceReport(
    DateTimeOffset CapturedAtUtc,
    string Commit,
    bool TrackedSourceDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    int Repeat,
    IReadOnlyList<HtmlOwnedDocumentEvidenceMeasurement> Measurements,
    IReadOnlyList<string> Failures);

internal sealed class HtmlOwnedDocumentBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<HtmlOwnedDocumentBudget> Budgets { get; set; } = new();
}

internal sealed class HtmlOwnedDocumentBudget {
    public string Operation { get; set; } = string.Empty;
    public string Scale { get; set; } = string.Empty;
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
    public long MaxRetainedManagedHeapGrowthBytes { get; set; }
    public long MaxPeakManagedHeapGrowthBytes { get; set; }
    public long MaxAbsoluteProcessPeakWorkingSetBytes { get; set; }
    public int MaxOutputCharacters { get; set; }
}
