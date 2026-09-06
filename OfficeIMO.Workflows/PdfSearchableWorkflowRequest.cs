using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Workflows;

/// <summary>Creates a searchable PDF through a caller-owned OCR engine and guarded artifact publication.</summary>
public sealed class PdfSearchableWorkflowRequest {
    /// <summary>Caller-provided identifier.</summary>
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    /// <summary>Source filesystem path or provider location.</summary>
    public string InputPath { get; set; } = string.Empty;
    /// <summary>Optional reopenable provider input.</summary>
    public OfficeWorkflowStreamInput? InputStream { get; set; }
    /// <summary>Explicit output filesystem path or provider location.</summary>
    public string OutputPath { get; set; } = string.Empty;
    /// <summary>Optional provider output, requiring confirmed direct-write consent and Replace policy.</summary>
    public OfficeWorkflowStreamOutput? OutputStream { get; set; }
    /// <summary>Existing destination policy.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; }
    /// <summary>Host authorization checked immediately before publication.</summary>
    public IOfficeWorkflowPublicationGuard? PublicationGuard { get; set; }
    /// <summary>Input and output byte limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
    /// <summary>Password used to open an encrypted source.</summary>
    public string? PdfPassword { get; set; }
    /// <summary>Recognition, page selection, rendering, and confidence options captured before asynchronous work.</summary>
    public PdfOcrMergeOptions Ocr { get; set; } = new();
    /// <summary>Optional review before mutation. Return eligible word instances from the supplied review.
    /// The runner validates the selection and rechecks source identity before publication.</summary>
    /// <remarks>The callback must honor cancellation. It does not grant permission to publish or bypass
    /// the output conflict policy or publication guard. Without a callback, all eligible words are used.</remarks>
    public Func<PdfSearchableOcrReview, CancellationToken, Task<IReadOnlyList<PdfRecognizedWord>>>? ReviewAsync { get; set; }
}

/// <summary>Searchable-PDF outcome; recognition metadata describes the generated artifact.</summary>
public sealed class PdfSearchableWorkflowResult {
    internal PdfSearchableWorkflowResult(OfficeWorkflowStatus status, string? outputPath, string summary,
        int addedWordCount, IReadOnlyList<int> modifiedPages, string? provider,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, OfficeWorkflowOutputRecovery? recovery = null) {
        Status = status;
        OutputPath = outputPath;
        Summary = summary;
        AddedWordCount = addedWordCount;
        ModifiedPages = Array.AsReadOnly(modifiedPages.ToArray());
        Provider = provider;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
        Recovery = recovery;
    }
    /// <summary>Terminal publication state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Verified published destination, present only on success.</summary>
    public string? OutputPath { get; }
    /// <summary>User-facing outcome or failure.</summary>
    public string Summary { get; }
    /// <summary>Words added to the generated artifact.</summary>
    public int AddedWordCount { get; }
    /// <summary>Pages receiving searchable text.</summary>
    public IReadOnlyList<int> ModifiedPages { get; }
    /// <summary>Provider reported by recognition.</summary>
    public string? Provider { get; }
    /// <summary>Validation, publication, and cleanup diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Retained artifact when provider publication needs attention.</summary>
    public OfficeWorkflowOutputRecovery? Recovery { get; }
}
