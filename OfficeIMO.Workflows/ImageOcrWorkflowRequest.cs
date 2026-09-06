using OfficeIMO.Reader;

namespace OfficeIMO.Workflows;

/// <summary>Recognizes a standalone image and publishes reviewed text as UTF-8.</summary>
public sealed class ImageOcrWorkflowRequest {
    /// <summary>Source filesystem path or provider location.</summary>
    public string InputPath { get; set; } = string.Empty;
    /// <summary>Optional reopenable provider input.</summary>
    public OfficeWorkflowStreamInput? InputStream { get; set; }
    /// <summary>Explicit text output path or provider location, with a .txt filename.</summary>
    public string OutputPath { get; set; } = string.Empty;
    /// <summary>Optional provider output, requiring explicit direct-write consent.</summary>
    public OfficeWorkflowStreamOutput? OutputStream { get; set; }
    /// <summary>Existing destination policy.</summary>
    public OfficeWorkflowConflictPolicy ConflictPolicy { get; set; }
    /// <summary>Host authorization checked immediately before publication.</summary>
    public IOfficeWorkflowPublicationGuard? PublicationGuard { get; set; }
    /// <summary>Input and output byte limits.</summary>
    public OfficeWorkflowLimits Limits { get; set; } = new();
    /// <summary>Bounded Reader OCR configuration, captured before asynchronous work.</summary>
    public OfficeDocumentOcrExecutionOptions Ocr { get; set; } = new();
    /// <summary>Optional review and correction of recognized text before publication.
    /// Cancellation aborts publication even if the callback fails to finish promptly.</summary>
    public Func<ImageOcrWorkflowReview, CancellationToken, Task<string>>? ReviewAsync { get; set; }
}

/// <summary>Image snapshot and Reader OCR evidence presented before text publication.</summary>
public sealed class ImageOcrWorkflowReview {
    private readonly byte[] _image;
    internal ImageOcrWorkflowReview(string name, byte[] image, OfficeDocumentOcrExecutionResult recognition, string text) {
        SourceName = name;
        _image = (byte[])image.Clone();
        Recognition = recognition;
        Text = text;
    }
    /// <summary>Original source filename.</summary>
    public string SourceName { get; }
    /// <summary>Recognized text before user corrections.</summary>
    public string Text { get; }
    /// <summary>Bounded recognition evidence, including word geometry and provider diagnostics.
    /// Changing this evidence does not change the source snapshot or published text.</summary>
    public OfficeDocumentOcrExecutionResult Recognition { get; }
    /// <summary>Returns an independent copy of the source image for preview.</summary>
    public byte[] GetImageBytes() => (byte[])_image.Clone();
}

/// <summary>Terminal image OCR publication result.</summary>
public sealed class ImageOcrWorkflowResult {
    internal ImageOcrWorkflowResult(OfficeWorkflowStatus status, string? output, string summary,
        int characters, IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics, OfficeWorkflowOutputRecovery? recovery = null) {
        Status = status; OutputPath = output; Summary = summary; CharacterCount = characters;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray()); Recovery = recovery;
    }
    /// <summary>Terminal publication state.</summary>
    public OfficeWorkflowStatus Status { get; }
    /// <summary>Verified published output, present only on success.</summary>
    public string? OutputPath { get; }
    /// <summary>User-facing outcome.</summary>
    public string Summary { get; }
    /// <summary>UTF-16 characters in the reviewed text prepared for publication.</summary>
    public int CharacterCount { get; }
    /// <summary>Recognition, publication, and cleanup diagnostics.</summary>
    public IReadOnlyList<OfficeWorkflowDiagnostic> Diagnostics { get; }
    /// <summary>Retained artifact when provider publication needs attention.</summary>
    public OfficeWorkflowOutputRecovery? Recovery { get; }
}
