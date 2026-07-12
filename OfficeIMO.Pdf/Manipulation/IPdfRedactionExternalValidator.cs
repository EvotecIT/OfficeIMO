namespace OfficeIMO.Pdf;

/// <summary>Optional development-time validator seam for tools such as qpdf, mutool, Ghostscript, or organization-specific forensic checks.</summary>
public interface IPdfRedactionExternalValidator {
    /// <summary>Validates one redacted artifact without becoming a runtime dependency of OfficeIMO.Pdf.</summary>
    PdfRedactionExternalValidationResult Validate(byte[] redactedPdf);
}

/// <summary>Result returned by an optional external redaction validator.</summary>
public sealed class PdfRedactionExternalValidationResult {
    /// <summary>Creates a validator result.</summary>
    public PdfRedactionExternalValidationResult(string validatorName, bool isValid, string? diagnostic = null) { Guard.NotNullOrWhiteSpace(validatorName, nameof(validatorName)); ValidatorName = validatorName; IsValid = isValid; Diagnostic = diagnostic; }
    /// <summary>Stable validator/tool name.</summary>
    public string ValidatorName { get; }
    /// <summary>True when the external validation passed.</summary>
    public bool IsValid { get; }
    /// <summary>Human-readable diagnostic.</summary>
    public string? Diagnostic { get; }
}
