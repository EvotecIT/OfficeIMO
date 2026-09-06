namespace OfficeIMO.Workflows;

/// <summary>One PDF or image OCR task in an explicitly ordered desktop or service session.</summary>
public sealed class OfficeOcrSessionRequest {
    /// <summary>Creates a searchable-PDF session item.</summary>
    public OfficeOcrSessionRequest(string id, PdfSearchableWorkflowRequest request) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A session item needs an id.", nameof(id));
        Id = id; Pdf = request ?? throw new ArgumentNullException(nameof(request));
    }
    /// <summary>Creates an image-to-text session item.</summary>
    public OfficeOcrSessionRequest(string id, ImageOcrWorkflowRequest request) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A session item needs an id.", nameof(id));
        Id = id; Image = request ?? throw new ArgumentNullException(nameof(request));
    }
    /// <summary>Caller identifier, unique within one session.</summary>
    public string Id { get; }
    /// <summary>Searchable-PDF request, or null for an image task.</summary>
    public PdfSearchableWorkflowRequest? Pdf { get; }
    /// <summary>Image-to-text request, or null for a PDF task.</summary>
    public ImageOcrWorkflowRequest? Image { get; }
}

/// <summary>One terminal session outcome. Previously completed outputs remain committed after cancellation.</summary>
public sealed class OfficeOcrSessionResult {
    internal OfficeOcrSessionResult(string id, OfficeWorkflowStatus status, string? output, string summary,
        IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, OfficeWorkflowOutputRecovery? recovery = null) {
        Id = id; Status = status; OutputPath = output; Summary = summary;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray()); Recovery = recovery;
    }
    /// <summary>Matching request identifier.</summary>
    public string Id { get; }
    /// <summary>Terminal publication state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Verified output, present only on success.</summary>
    public string? OutputPath { get; }
    /// <summary>Result or actionable failure description.</summary>
    public string Summary { get; }
    /// <summary>Recognition, publication, and cleanup diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Retained provider output requiring attention.</summary>
    public OfficeWorkflowOutputRecovery? Recovery { get; }
}
