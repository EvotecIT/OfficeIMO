namespace OfficeIMO.Pdf;

/// <summary>Single diagnostic reported while inspecting a PDF.</summary>
public sealed class PdfDiagnosticFinding {
    internal PdfDiagnosticFinding(
        PdfDiagnosticSeverity severity,
        string code,
        string message,
        int? objectNumber = null,
        int? pageNumber = null,
        long? bytes = null) {
        Severity = severity;
        Code = code;
        Message = message;
        ObjectNumber = objectNumber;
        PageNumber = pageNumber;
        Bytes = bytes;
    }

    /// <summary>Finding severity.</summary>
    public PdfDiagnosticSeverity Severity { get; }

    /// <summary>Stable finding code.</summary>
    public string Code { get; }

    /// <summary>Human-readable finding message.</summary>
    public string Message { get; }

    /// <summary>Related PDF object number, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Related one-based page number, when known.</summary>
    public int? PageNumber { get; }

    /// <summary>Related byte size or estimated byte saving, when known.</summary>
    public long? Bytes { get; }
}
