namespace OfficeIMO.Pdf;

/// <summary>Single finding reported by the lightweight PDF signature validator.</summary>
public sealed class PdfSignatureValidationFinding {
    internal PdfSignatureValidationFinding(
        PdfDiagnosticSeverity severity,
        string code,
        string message,
        int? signatureObjectNumber = null,
        int? fieldObjectNumber = null,
        bool isCryptographic = false) {
        Severity = severity;
        Code = code;
        Message = message;
        SignatureObjectNumber = signatureObjectNumber;
        FieldObjectNumber = fieldObjectNumber;
        IsCryptographic = isCryptographic;
    }

    /// <summary>Finding severity.</summary>
    public PdfDiagnosticSeverity Severity { get; }

    /// <summary>Stable finding code.</summary>
    public string Code { get; }

    /// <summary>Human-readable finding message.</summary>
    public string Message { get; }

    /// <summary>Signature value dictionary object number, when known.</summary>
    public int? SignatureObjectNumber { get; }

    /// <summary>Owning AcroForm signature field object number, when known.</summary>
    public int? FieldObjectNumber { get; }

    /// <summary>True when the finding came from an optional cryptography provider rather than PDF structure analysis.</summary>
    public bool IsCryptographic { get; }
}
