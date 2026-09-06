namespace OfficeIMO.Workflows;

/// <summary>Splits a PDF into consecutive parts and publishes validated outputs.</summary>
public sealed class PdfSplitWorkflowRequest {
    /// <summary>Caller-provided identifier.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    /// <summary>Original source location.</summary>
    public required string InputPath { get; set; }
    /// <summary>Optional reopenable provider access.</summary>
    public OfficeWorkflowStreamInput? InputStream { get; set; }
    /// <summary>Destination folder. Local folders publish as a unit.</summary>
    public required string OutputDirectory { get; set; }
    /// <summary>Optional provider access. Requires Replace; verified parts survive a later failure.</summary>
    public OfficeWorkflowDirectoryOutput? DirectoryOutput { get; set; }
    /// <summary>Maximum pages in each consecutive part.</summary>
    public int PagesPerDocument { get; set; } = 1;
    /// <summary>Maximum permitted number of generated parts.</summary>
    public int MaximumParts { get; set; } = 1000;
    /// <summary>Optional input password.</summary>
    public string? PdfPassword { get; set; }
    /// <summary>Conflict policy for the local output folder or provider children.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; } = OfficeWorkflowConflictPolicy.Rename;
    /// <summary>Host authorization checked before publication.</summary>
    public IOfficeWorkflowPublicationGuard? PublicationGuard { get; set; }
    /// <summary>Input and aggregate output byte limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
}

/// <summary>One verified split output and its source page range.</summary>
public sealed record PdfSplitFile(
    string Path, int FirstSourcePage, int PageCount, long SizeBytes);

/// <summary>Split completion, verified files, and recovery records for uncertain provider writes.</summary>
public sealed class PdfSplitWorkflowResult {
    internal PdfSplitWorkflowResult(string id, OfficeWorkflowStatus status, string summary,
        IEnumerable<PdfSplitFile> files, IEnumerable<OfficeWorkflowDiagnostic> diagnostics,
        IEnumerable<OfficeWorkflowOutputRecovery>? recoveries = null) {
        Id = id; Status = status; Summary = summary;
        Files = Array.AsReadOnly(files.ToArray()); Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        OutputRecoveries = Array.AsReadOnly(recoveries?.ToArray() ?? []);
    }
    /// <summary>Request identifier.</summary>
    public string Id { get; }
    /// <summary>Terminal publication status.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Whether all parts were published and verified.</summary>
    public bool Succeeded => Status == OfficeWorkflowStatus.Completed;
    /// <summary>Human-readable completion or failure description.</summary>
    public string Summary { get; }
    /// <summary>Successfully published files, including partial completion.</summary>
    public IReadOnlyList<PdfSplitFile> Files { get; }
    /// <summary>Validation and publication diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Retained recovery artifacts for outputs requiring attention.</summary>
    public IReadOnlyList<OfficeWorkflowOutputRecovery> OutputRecoveries { get; }
}
