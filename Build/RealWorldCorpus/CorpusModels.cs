using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Reader;

namespace OfficeIMO.RealWorldCorpus;

internal static class CorpusOutcomes {
    public const string Classification = "classification";
    public const string Probe = "probe";
    public const string NotEligible = "not-eligible";
    public const string NotSelected = "not-selected";
    public const string Duplicate = "duplicate";
    public const string SkippedOversize = "skipped-oversize";
    public const string ClassificationFailed = "classification-failed";
    public const string ClassificationTimedOut = "classification-timed-out";
    public const string Completed = "completed";
    public const string CompletedWithWarnings = "completed-with-warnings";
    public const string CompletedWithErrors = "completed-with-errors";
    public const string Rejected = "rejected-by-policy";
    public const string Failed = "failed";
    public const string TimedOut = "timed-out";
}

internal static class CorpusJson {
    public static JsonSerializerOptions Options { get; } = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
    };
}

internal sealed class CorpusRunOptions {
    public string InputDirectory { get; set; } = string.Empty;
    public string JsonReportPath { get; set; } = string.Empty;
    public string MarkdownReportPath { get; set; } = string.Empty;
    public string CorpusId { get; set; } = "local-corpus";
    public string? SourceUri { get; set; }
    public string? ArchiveSha256 { get; set; }
    public int MaxPerFormat { get; set; } = 100;
    public int MaxTotal { get; set; } = 600;
    public long MaxFileBytes { get; set; } = 50L * 1024L * 1024L;
    public int MaxTraversalEntries { get; set; } = 5_000;
    public int TimeoutSeconds { get; set; } = 30;
    public int Parallelism { get; set; } = 4;
    public bool IncludeSourceNames { get; set; }
    public IReadOnlyList<ReaderInputKind> Formats { get; set; } = Array.Empty<ReaderInputKind>();
}

internal sealed class CorpusWorkerOptions {
    public string InputPath { get; set; } = string.Empty;
    public long MaxFileBytes { get; set; }
    public string? ExpectedSha256 { get; set; }
    public string Stage { get; set; } = string.Empty;
}

internal sealed class CorpusWorkerResult {
    public string Stage { get; set; } = string.Empty;
    public bool Succeeded { get; set; }
    public string? Sha256 { get; set; }
    public ReaderInputKind ExtensionKind { get; set; }
    public ReaderInputKind ContentKind { get; set; }
    public ReaderDetectionConfidence ContentConfidence { get; set; }
    public ReaderInputKind DetectedKind { get; set; }
    public ReaderDetectionConfidence Confidence { get; set; }
    public bool IsMismatch { get; set; }
    public IReadOnlyList<string> Evidence { get; set; } = Array.Empty<string>();
    public int ChunkCount { get; set; }
    public int PageCount { get; set; }
    public int BlockCount { get; set; }
    public int AssetCount { get; set; }
    public int InformationDiagnostics { get; set; }
    public int WarningDiagnostics { get; set; }
    public int ErrorDiagnostics { get; set; }
    public IReadOnlyList<string> DiagnosticCodes { get; set; } = Array.Empty<string>();
    public string? ExceptionType { get; set; }

    public static CorpusWorkerResult Failure(string stage, Exception exception) => new() {
        Stage = stage,
        Succeeded = false,
        ExceptionType = exception.GetType().FullName ?? exception.GetType().Name
    };
}

internal sealed class CorpusReport {
    public int SchemaVersion { get; set; } = 1;
    public string MeasurementStatus { get; set; } = "measured";
    public DateTimeOffset StartedUtc { get; set; }
    public DateTimeOffset CompletedUtc { get; set; }
    public CorpusProvenance Provenance { get; set; } = new();
    public CorpusConfiguration Configuration { get; set; } = new();
    public CorpusEnvironment Environment { get; set; } = new();
    public CorpusTotals Totals { get; set; } = new();
    public IReadOnlyList<CorpusStratum> Strata { get; set; } = Array.Empty<CorpusStratum>();
    public IReadOnlyList<CorpusFileRecord> Files { get; set; } = Array.Empty<CorpusFileRecord>();
}

internal sealed class CorpusProvenance {
    public string CorpusId { get; set; } = string.Empty;
    public string? SourceUri { get; set; }
    public string? ArchiveSha256 { get; set; }
}

internal sealed class CorpusConfiguration {
    public IReadOnlyList<ReaderInputKind> Formats { get; set; } = Array.Empty<ReaderInputKind>();
    public int MaxPerFormat { get; set; }
    public int MaxTotal { get; set; }
    public long MaxFileBytes { get; set; }
    public int MaxTraversalEntries { get; set; }
    public int TimeoutSeconds { get; set; }
    public int Parallelism { get; set; }
    public bool SourceNamesIncluded { get; set; }
    public CorpusReaderPolicyConfiguration ReaderPolicy { get; set; } = new();
    public CorpusPackagePolicyConfiguration PackagePolicy { get; set; } = new();
    public string Selection { get; set; } = "sha256-ascending-stratified";
    public string Operation { get; set; } = "officeimo-reader-normalized-read";
}

internal sealed class CorpusReaderPolicyConfiguration {
    public ReaderDetectionMode DetectionMode { get; set; }
    public bool InspectContainers { get; set; }
    public int DetectionMaxProbeBytes { get; set; }
    public int DetectionMaxContainerEntries { get; set; }
    public int ReadMaxCharacters { get; set; }
    public int ReadMaxTableRows { get; set; }
    public bool ComputeHashes { get; set; }
}

internal sealed class CorpusPackagePolicyConfiguration {
    public long MaxPackageBytes { get; set; }
    public int MaxPartCount { get; set; }
    public long MaxPartUncompressedBytes { get; set; }
    public long MaxXmlCharactersInPart { get; set; }
    public long MaxTotalUncompressedBytes { get; set; }
    public double MaxCompressionRatio { get; set; }
}

internal sealed class CorpusEnvironment {
    public string Framework { get; set; } = string.Empty;
    public string OperatingSystem { get; set; } = string.Empty;
    public string ProcessArchitecture { get; set; } = string.Empty;
}

internal sealed class CorpusTotals {
    public int Discovered { get; set; }
    public int Oversize { get; set; }
    public int ClassificationFailed { get; set; }
    public int ClassificationTimedOut { get; set; }
    public int DuplicateContent { get; set; }
    public int EligibleUnique { get; set; }
    public int Selected { get; set; }
    public int Completed { get; set; }
    public int CompletedWithWarnings { get; set; }
    public int CompletedWithErrors { get; set; }
    public int RejectedByPolicy { get; set; }
    public int Failed { get; set; }
    public int TimedOut { get; set; }
}

internal sealed class CorpusStratum {
    public ReaderInputKind Format { get; set; }
    public int EligibleUnique { get; set; }
    public int RequestedMaximum { get; set; }
    public int Selected { get; set; }
    public int Completed { get; set; }
    public int CompletedWithWarnings { get; set; }
    public int CompletedWithErrors { get; set; }
    public int RejectedByPolicy { get; set; }
    public int Failed { get; set; }
    public int TimedOut { get; set; }
    public bool CorpusUnderfilled { get; set; }
}

internal sealed class CorpusFileRecord {
    [JsonIgnore]
    public string FullPath { get; set; } = string.Empty;
    public int InventoryIndex { get; set; }
    public string? SourceName { get; set; }
    public string Extension { get; set; } = string.Empty;
    public long SizeBytes { get; set; }
    public string? Sha256 { get; set; }
    public ReaderInputKind ExtensionKind { get; set; }
    public ReaderInputKind ContentKind { get; set; }
    public ReaderDetectionConfidence ContentConfidence { get; set; }
    public ReaderInputKind DetectedKind { get; set; }
    public ReaderDetectionConfidence Confidence { get; set; }
    public bool IsMismatch { get; set; }
    public IReadOnlyList<string> DetectionEvidence { get; set; } = Array.Empty<string>();
    public bool Selected { get; set; }
    public string Outcome { get; set; } = CorpusOutcomes.NotEligible;
    public string? FailureStage { get; set; }
    public long? ClassificationDurationMilliseconds { get; set; }
    public long? ProbeDurationMilliseconds { get; set; }
    public int? ChunkCount { get; set; }
    public int? PageCount { get; set; }
    public int? BlockCount { get; set; }
    public int? AssetCount { get; set; }
    public int InformationDiagnostics { get; set; }
    public int WarningDiagnostics { get; set; }
    public int ErrorDiagnostics { get; set; }
    public IReadOnlyList<string> DiagnosticCodes { get; set; } = Array.Empty<string>();
    public string? ExceptionType { get; set; }
}
