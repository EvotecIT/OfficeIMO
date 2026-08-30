using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.PdfQualityCorpus;

internal static class QualityJson {
    internal static JsonSerializerOptions Options { get; } = new() {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}

internal sealed class QualityManifest {
    public int Version { get; set; }
    public string Authority { get; set; } = string.Empty;
    public IReadOnlyList<QualitySource> Sources { get; set; } = Array.Empty<QualitySource>();
    public IReadOnlyList<QualityCase> Cases { get; set; } = Array.Empty<QualityCase>();
}

internal sealed class QualitySource {
    public string Id { get; set; } = string.Empty;
    public string Repository { get; set; } = string.Empty;
    public string Commit { get; set; } = string.Empty;
    public string License { get; set; } = string.Empty;
}

internal sealed class QualityCase {
    public string Id { get; set; } = string.Empty;
    public string File { get; set; } = string.Empty;
    public string Source { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string Sha256 { get; set; } = string.Empty;
    public long ByteLength { get; set; }
    public int PageCount { get; set; }
    public int MinimumTextCharacters { get; set; }
    public int MinimumAttachments { get; set; }
    public int MinimumLinks { get; set; }
    public int MinimumAnnotations { get; set; }
    public int? MinimumOptionalContentGroups { get; set; }
    public int? MinimumFonts { get; set; }
    public int? MinimumEmbeddedFonts { get; set; }
    public int? MinimumSubsetFonts { get; set; }
    public int? MaximumMissingToUnicodeFonts { get; set; }
    public IReadOnlyList<string> ExpectedAnnotationActionTypes { get; set; } = Array.Empty<string>();
    public IReadOnlyList<string> ExpectedRepairCodes { get; set; } = Array.Empty<string>();
    public bool ExpectedRenderSucceeded { get; set; }
    public IReadOnlyList<string> ExpectedRenderDiagnosticCodes { get; set; } = Array.Empty<string>();
    public string ExpectedMutationMode { get; set; } = string.Empty;
    public IReadOnlyList<string> Features { get; set; } = Array.Empty<string>();
}

internal sealed class QualityRunOptions {
    public string ManifestPath { get; set; } = string.Empty;
    public string RootDirectory { get; set; } = string.Empty;
    public string JsonReportPath { get; set; } = string.Empty;
    public string MarkdownReportPath { get; set; } = string.Empty;
    public long MaxFileBytes { get; set; } = 128L * 1024L * 1024L;
    public int MaxRenderPages { get; set; } = 4;
    public int TimeoutSeconds { get; set; } = 60;
    public int Parallelism { get; set; } = Math.Max(1, Math.Min(4, Environment.ProcessorCount));
    public long MaxWorkerMemoryBytes { get; set; } = 1024L * 1024L * 1024L;
}

internal sealed class QualityReportConfiguration {
    public string ManifestFileName { get; set; } = string.Empty;
    public string ManifestSha256 { get; set; } = string.Empty;
    public long MaxFileBytes { get; set; }
    public int MaxRenderPages { get; set; }
    public int TimeoutSeconds { get; set; }
    public int Parallelism { get; set; }
    public long MaxWorkerMemoryBytes { get; set; }
}

internal sealed class QualityProbeOptions {
    public string ManifestPath { get; set; } = string.Empty;
    public string RootDirectory { get; set; } = string.Empty;
    public string CaseId { get; set; } = string.Empty;
    public long MaxFileBytes { get; set; }
    public int MaxRenderPages { get; set; }
}

internal sealed class QualityCheckResult {
    public string Name { get; set; } = string.Empty;
    public bool Succeeded { get; set; }
    public long DurationMilliseconds { get; set; }
    public string? ExceptionType { get; set; }
    public string? Message { get; set; }
}

internal sealed class QualityExpectationResult {
    public string Name { get; set; } = string.Empty;
    public bool Succeeded { get; set; }
    public string Expected { get; set; } = string.Empty;
    public string Actual { get; set; } = string.Empty;
}

internal sealed class QualityCaseMetrics {
    public int PageCount { get; set; }
    public int TextCharacters { get; set; }
    public int ParagraphCount { get; set; }
    public int CrossPageParagraphCount { get; set; }
    public int TableCount { get; set; }
    public int CrossPageTableCount { get; set; }
    public int FontCount { get; set; }
    public int EmbeddedFontCount { get; set; }
    public int SubsetFontCount { get; set; }
    public int MissingToUnicodeFontCount { get; set; }
    public int UnreadableToUnicodeFontCount { get; set; }
    public int FontResourceReferenceCount { get; set; }
    public int ImageCount { get; set; }
    public int ImagePlacementCount { get; set; }
    public int AttachmentCount { get; set; }
    public int LinkCount { get; set; }
    public int AnnotationCount { get; set; }
    public int OptionalContentGroupCount { get; set; }
    public int RenderAttemptedPages { get; set; }
    public int RenderSucceededPages { get; set; }
    public long RenderOutputBytes { get; set; }
    public IReadOnlyList<string> RenderDiagnosticCodes { get; set; } = Array.Empty<string>();
    public IReadOnlyList<string> RepairCodes { get; set; } = Array.Empty<string>();
    public IReadOnlyList<string> AnnotationActionTypes { get; set; } = Array.Empty<string>();
    public string MutationMode { get; set; } = string.Empty;
    public IReadOnlyList<string> MutationBlockerCodes { get; set; } = Array.Empty<string>();
    public int MutationPlanCount { get; set; }
    public int FullRewriteMutationPlanCount { get; set; }
    public int AppendOnlyMutationPlanCount { get; set; }
    public int BlockedMutationPlanCount { get; set; }
    public IReadOnlyDictionary<string, string> MutationPlanModes { get; set; } = new Dictionary<string, string>();
    public int DeclaredComplianceClaimCount { get; set; }
    public int RecognizedComplianceClaimCount { get; set; }
    public int ClaimableComplianceClaimCount { get; set; }
    public int UnsupportedComplianceClaimCount { get; set; }
    public IReadOnlyList<string> DeclaredComplianceClaimStatuses { get; set; } = Array.Empty<string>();
}

internal sealed class QualityCaseResult {
    public string Id { get; set; } = string.Empty;
    public string SourceId { get; set; } = string.Empty;
    public string SourcePath { get; set; } = string.Empty;
    public string Sha256 { get; set; } = string.Empty;
    public long ByteLength { get; set; }
    public IReadOnlyList<string> Features { get; set; } = Array.Empty<string>();
    public string Outcome { get; set; } = "failed";
    public bool TimedOut { get; set; }
    public string? FailureCode { get; set; }
    public QualityCaseMetrics Metrics { get; set; } = new();
    public IReadOnlyList<QualityCheckResult> Checks { get; set; } = Array.Empty<QualityCheckResult>();
    public IReadOnlyList<QualityExpectationResult> Expectations { get; set; } = Array.Empty<QualityExpectationResult>();
    public long DurationMilliseconds { get; set; }
    public long WorkerWallClockMilliseconds { get; set; }
    public long WorkerCpuMilliseconds { get; set; }
    public long PeakWorkingSetBytes { get; set; }
    public double OperationalScore => Checks.Count == 0 ? 0D : (double)Checks.Count(check => check.Succeeded) / Checks.Count;
    public double ExpectationScore => Expectations.Count == 0 ? 1D : (double)Expectations.Count(expectation => expectation.Succeeded) / Expectations.Count;
}

internal sealed class QualityEnvironment {
    public string Framework { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public string ProcessArchitecture { get; set; } = string.Empty;
    public string EngineAssemblyVersion { get; set; } = string.Empty;
}

internal sealed class QualityTotals {
    public int Cases { get; set; }
    public int Passed { get; set; }
    public int Failed { get; set; }
    public int TimedOut { get; set; }
    public int OperationalChecks { get; set; }
    public int OperationalChecksPassed { get; set; }
    public int Expectations { get; set; }
    public int ExpectationsPassed { get; set; }
    public long InputBytes { get; set; }
    public int Pages { get; set; }
    public long DurationMilliseconds { get; set; }
    public long PeakWorkingSetBytes { get; set; }
    public double OperationalScore => OperationalChecks == 0 ? 0D : (double)OperationalChecksPassed / OperationalChecks;
    public double ExpectationScore => Expectations == 0 ? 1D : (double)ExpectationsPassed / Expectations;
}

internal sealed class QualityReport {
    public int SchemaVersion { get; set; } = 1;
    public string MeasurementStatus { get; set; } = "measured";
    public DateTimeOffset StartedUtc { get; set; }
    public DateTimeOffset CompletedUtc { get; set; }
    public string Authority { get; set; } = string.Empty;
    public IReadOnlyList<QualitySource> Sources { get; set; } = Array.Empty<QualitySource>();
    public QualityEnvironment Environment { get; set; } = new();
    public QualityReportConfiguration Configuration { get; set; } = new();
    public QualityTotals Totals { get; set; } = new();
    public IReadOnlyList<QualityCaseResult> Cases { get; set; } = Array.Empty<QualityCaseResult>();
}
