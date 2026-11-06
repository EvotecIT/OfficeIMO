namespace OfficeIMO.Pdf;

/// <summary>Provider-owned cryptographic validation finding.</summary>
public sealed class PdfSignatureCryptographicFinding {
    /// <summary>Creates a provider finding with a stable code and human-readable message.</summary>
    public PdfSignatureCryptographicFinding(PdfDiagnosticSeverity severity, string code, string message) {
        ArgumentException.ThrowIfNullOrWhiteSpace(code);
        ArgumentException.ThrowIfNullOrWhiteSpace(message);
        Severity = severity;
        Code = code;
        Message = message;
    }

    /// <summary>Finding severity.</summary>
    public PdfDiagnosticSeverity Severity { get; }

    /// <summary>Stable provider finding code.</summary>
    public string Code { get; }

    /// <summary>Human-readable finding message.</summary>
    public string Message { get; }
}
