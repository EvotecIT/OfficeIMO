using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

/// <summary>Incremental file conversion through the existing workflow routes. Checkpoints are optional.</summary>
public sealed record OfficeConversionBatchRequest {
    /// <summary>Source directory; descendants are discovered incrementally without following links.</summary>
    public string? InputDirectory { get; set; }
    /// <summary>Explicit local files. When InputDirectory is set, files must be inside that root and retain relative paths.</summary>
    public string[]? InputPaths { get; set; }
    /// <summary>Destination extension from the configured runner's executable routes, including its leading dot.</summary>
    public string TargetExtension { get; set; } = ".pdf";
    /// <summary>Optional explicit route for ambiguous inputs; otherwise the configured runner's routes select the conversion.</summary>
    public string? ConversionRouteId { get; set; }
    /// <summary>Runtime PDF source password, equivalent to OfficeWorkflowRequest.PdfPassword. Not stored in checkpoints.</summary>
    public string? PdfPassword { get; set; }
    /// <summary>Existing typed settings. Each selected route receives only its applicable settings.</summary>
    public OfficeWorkflowConversionOptions ConversionOptions { get; set; } = new();
    /// <summary>Cross-format output intent; format-specific settings take precedence when supplied.</summary>
    public OfficeWorkflowOutputProfile OutputProfile { get; set; } = OfficeWorkflowOutputProfile.Faithful;
    /// <summary>Existing publication conflict policy. Durable checkpoint jobs require Fail.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Fail;
    /// <summary>Discover descendants when selecting a directory.</summary>
    public bool Recursive { get; set; } = true;
    /// <summary>Optional source extensions to select during directory discovery. Unselected files are reported as skipped.</summary>
    public string[]? SourceExtensions { get; set; }
    /// <summary>Separate destination directory. Relative paths retain the full source name plus the target extension.</summary>
    public required string OutputDirectory { get; set; }
    /// <summary>Private durable state directory, outside both source and output trees. Only built-in routes support checkpoints.</summary>
    public string? CheckpointDirectory { get; set; }
    /// <summary>Maximum simultaneously executing files, from 1 through 32.</summary>
    public int MaximumConcurrency { get; set; } = 2;
    /// <summary>Maximum discovered files in one run, including skipped files. Folder discovery is incremental.</summary>
    public int MaximumFiles { get; set; } = 1_000_000;
    /// <summary>Maximum input bytes per document.</summary>
    public long MaximumInputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum output bytes per document.</summary>
    public long MaximumOutputBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum XML characters per part for built-in Open XML, Draw and Visio conversions, or in a complete legacy Visio XML source.</summary>
    public long MaximumXmlCharactersInPart { get; set; } = 10L * 1024 * 1024;
    /// <summary>Retry recorded failures. Completed files still require source and output hash verification.</summary>
    public bool RetryFailed { get; set; }
}

/// <summary>One file outcome. Skipped files have no execution status or output. Callbacks may arrive concurrently.</summary>
public sealed record OfficeConversionBatchItemResult(string InputPath, string? OutputPath, OfficeWorkflowStatus? Status,
    bool Reused, string Summary, IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics, bool Skipped = false);

/// <summary>Bounded summary. Completed outputs survive cancellation and item failures.</summary>
public sealed record OfficeConversionBatchResult(long Selected, long Completed, long Reused, long Failed, bool Cancelled, long Skipped = 0);

internal sealed record OfficeConversionBatchPlan(string Schema, string ConfigurationSha256);
internal sealed record OfficeConversionBatchReceipt(string Schema, string ConfigurationSha256, string InputSha256,
    string? OutputSha256, OfficeWorkflowStatus Status, string Summary, OfficeConversionBatchStoredDiagnostic[] Diagnostics, int DiagnosticCount,
    string? PendingStageId = null);
internal sealed record OfficeConversionBatchStoredDiagnostic(string Code, string Message, OfficeWorkflowDiagnosticSeverity Severity,
    string? Stage, string? Source, string? LossKind, string? Location = null);

[JsonSourceGenerationOptions(UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeConversionBatchResult))]
[JsonSerializable(typeof(OfficeConversionBatchPlan))]
[JsonSerializable(typeof(OfficeConversionBatchReceipt))]
internal partial class OfficeConversionBatchJsonContext : JsonSerializerContext { }
