using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.PdfCorpusReliability;

/// <summary>Outcome of one operation, or of a whole file when the worker process was lost.</summary>
internal static class Outcome {
    public const string Ok = "ok";
    /// <summary>A documented OfficeIMO rejection of bad or unsupported input.</summary>
    public const string Rejected = "rejected";
    /// <summary>A rewrite refused by the preservation policy (PdfMutationBlockedException), not an input defect.</summary>
    public const string Blocked = "blocked";
    /// <summary>An exception type that signals a defect rather than a reasoned rejection.</summary>
    public const string Unexpected = "unexpected";
    /// <summary>The operation completed but its result disagrees with the source (round-trip loss).</summary>
    public const string Mismatch = "mismatch";
    public const string Skipped = "skipped";
    public const string Memory = "memory";
    public const string Timeout = "timeout";
    public const string Crash = "crash";

    public static int Severity(string outcome) => outcome switch {
        Ok => 0,
        Skipped => 0,
        Blocked => 0,
        Rejected => 1,
        Mismatch => 2,
        Unexpected => 3,
        Memory => 4,
        Timeout => 4,
        Crash => 5,
        _ => 3
    };

    /// <summary>Worst operation outcome; a file with no defect but a policy-blocked rewrite reports blocked.</summary>
    public static string Worst(IReadOnlyList<OperationResult> operations) {
        if (operations.Count == 0) return Crash;
        string worst = operations.MaxBy(static op => Severity(op.Outcome))!.Outcome;
        if (Severity(worst) > 0) return worst;
        return operations.Any(static op => op.Outcome == Blocked) ? Blocked : Ok;
    }
}

internal static class Operations {
    public const string Load = "load";
    public const string PageCount = "page-count";
    public const string Text = "text";
    public const string Read = "structured-read";
    public const string Render = "render";
    public const string Save = "save";
    public const string RoundTrip = "round-trip";
    public const string Rewrite = "rewrite-round-trip";
    public const string Repair = "repair-report";

    public static readonly string[] All = { Load, PageCount, Text, Read, Render, Save, RoundTrip, Rewrite, Repair };
}

internal sealed class CorpusEntry {
    public string Path { get; set; } = "";
    public string FullPath { get; set; } = "";
    public string Collection { get; set; } = "";
    public string Source { get; set; } = "";
    public string? Sha256 { get; set; }
    public long Bytes { get; set; }
    public int? QpdfExitCode { get; set; }
    public string[]? Passwords { get; set; }
    public bool PasswordsUnpublished { get; set; }
    public bool CertificateEncrypted { get; set; }
}

internal sealed class WorkerRequest {
    public string FullPath { get; set; } = "";
    public string[]? Passwords { get; set; }
    public int RenderPages { get; set; } = 3;
    public string? PngDirectory { get; set; }
}

internal sealed class OperationResult {
    public string Name { get; set; } = "";
    public string Outcome { get; set; } = "";
    public long Milliseconds { get; set; }
    public string? ExceptionType { get; set; }
    public string? Message { get; set; }
    public string? Frame { get; set; }
    public string? Detail { get; set; }
}

internal sealed class WorkerResponse {
    public string? Sha256 { get; set; }
    public long Bytes { get; set; }
    public int? PageCount { get; set; }
    public bool UsedPassword { get; set; }
    public long TextLength { get; set; }
    public int? Paragraphs { get; set; }
    public int? Tables { get; set; }
    public int RenderedPages { get; set; }
    public int RenderAttemptedPages { get; set; }
    public long? SavedBytes { get; set; }
    public long? RewrittenBytes { get; set; }
    public int? RepairCount { get; set; }
    public int? DetectedOnlyCount { get; set; }
    public string[]? RepairCodes { get; set; }
    public List<OperationResult> Operations { get; set; } = new();
}

internal sealed class FileResult {
    public string Path { get; set; } = "";
    public string Collection { get; set; } = "";
    public string Source { get; set; } = "";
    public string? Sha256 { get; set; }
    public bool Sha256Verified { get; set; }
    public long Bytes { get; set; }
    public int? QpdfExitCode { get; set; }
    public string? EncryptionNote { get; set; }
    /// <summary>Worst outcome across operations, or the process-level outcome.</summary>
    public string Outcome { get; set; } = "";
    public string? ProcessFailure { get; set; }
    public string? InFlightOperation { get; set; }
    public long WallMilliseconds { get; set; }
    public long PeakWorkingSetBytes { get; set; }
    public WorkerResponse? Worker { get; set; }
}

internal static class ReliabilityJson {
    public static readonly JsonSerializerOptions Line = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    public static readonly JsonSerializerOptions Indented = new(Line) { WriteIndented = true };
}
