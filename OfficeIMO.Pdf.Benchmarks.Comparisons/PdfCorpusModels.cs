using System.Text.Json.Serialization;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed class PdfCorpusManifest {
    public int SchemaVersion { get; init; }
    public string Description { get; init; } = string.Empty;
    public List<PdfCorpusEntry> Entries { get; init; } = new();
}

internal sealed class PdfCorpusEntry {
    public string Id { get; init; } = string.Empty;
    public string SourceKind { get; init; } = string.Empty;
    public string? SourcePath { get; init; }
    public string? Url { get; init; }
    public string? Sha256 { get; init; }
    public string? Generator { get; init; }
    public string Producer { get; init; } = string.Empty;
    public string License { get; init; } = string.Empty;
    public string? LicenseUrl { get; init; }
    public string Tier { get; init; } = string.Empty;
    public int? ExpectedPages { get; init; }
    public double MinimumTokenRecall { get; init; } = 0.75D;
    public List<string> Features { get; init; } = new();
    public List<string> RequiredText { get; init; } = new();
}

internal sealed record PdfCorpusReadResult(
    bool Success,
    string Oracle,
    int PageCount,
    int ExtractedCharacters,
    int OracleCharacters,
    double TokenRecall,
    string? Error,
    double ElapsedMilliseconds = 0D,
    long AllocatedBytes = 0L);

internal sealed record PdfCorpusManipulationResult(
    bool Success,
    string Status,
    int SelectedPages,
    int SplitDocuments,
    int MergedPages,
    IReadOnlyList<string> BlockerCodes,
    string? Error) {
    [JsonIgnore]
    internal bool IsFailure => Status is "Failed" or "NotRun";
}

internal sealed record PdfCorpusResult(
    string Id,
    string Producer,
    string Tier,
    string SourceKind,
    string Path,
    long Bytes,
    string Sha256,
    IReadOnlyList<string> Features,
    PdfCorpusReadResult Read,
    PdfCorpusManipulationResult Manipulation);

internal sealed record PdfCorpusClassSummary(
    string Dimension,
    string Name,
    int Documents,
    int SuccessfulReads,
    int FailedReads,
    int BlockedManipulations,
    int FailedManipulations,
    int TotalPages,
    double TotalReadMilliseconds,
    long TotalAllocatedBytes,
    double ReadFailureRate,
    double MillisecondsPerPage,
    double AllocatedBytesPerPage);

internal sealed record PdfCorpusReport(
    int SchemaVersion,
    DateTimeOffset CreatedUtc,
    string Runtime,
    string OperatingSystem,
    IReadOnlyList<PdfCorpusResult> Results,
    IReadOnlyList<PdfCorpusClassSummary> Classes) {
    [JsonIgnore]
    internal bool Success => Results.All(static result => result.Read.Success && !result.Manipulation.IsFailure);
}
