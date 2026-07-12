namespace OfficeIMO.Pdf;

/// <summary>Provider-owned cryptographic validation finding.</summary>
public sealed class PdfSignatureCryptographicFinding {
    /// <summary>Creates a provider finding with a stable code and human-readable message.</summary>
    public PdfSignatureCryptographicFinding(PdfDiagnosticSeverity severity, string code, string message) {
        if (string.IsNullOrWhiteSpace(code)) throw new ArgumentException("Finding code cannot be empty.", nameof(code));
        if (string.IsNullOrWhiteSpace(message)) throw new ArgumentException("Finding message cannot be empty.", nameof(message));
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
