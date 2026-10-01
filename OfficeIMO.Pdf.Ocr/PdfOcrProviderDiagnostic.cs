using OfficeIMO.Ocr;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Bounded, immutable provider diagnostic associated with a recognized PDF page.</summary>
public sealed class PdfOcrProviderDiagnostic {
    internal PdfOcrProviderDiagnostic(OcrDiagnostic diagnostic) {
        Code = diagnostic.Code ?? string.Empty; Message = diagnostic.Message ?? string.Empty;
        Severity = diagnostic.Severity; IsRecoverable = diagnostic.IsRecoverable;
    }
    /// <summary>Provider diagnostic code.</summary>
    public string Code { get; }
    /// <summary>Provider diagnostic message.</summary>
    public string Message { get; }
    /// <summary>Original provider severity.</summary>
    public OcrDiagnosticSeverity Severity { get; }
    /// <summary>Whether the provider permits recognition to continue.</summary>
    public bool IsRecoverable { get; }
    internal string DisplayText => string.IsNullOrWhiteSpace(Code) ? Message : Code + ": " + Message;
}
