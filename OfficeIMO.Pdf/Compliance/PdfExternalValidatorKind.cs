namespace OfficeIMO.Pdf;

/// <summary>
/// External validator family used as formal compliance evidence.
/// </summary>
public enum PdfExternalValidatorKind {
    /// <summary>veraPDF archival-profile validator, commonly used for PDF/A validation.</summary>
    VeraPdf,

    /// <summary>PDF/UA accessibility validator.</summary>
    PdfUaValidator,

    /// <summary>Mustang validator, commonly used for Factur-X and ZUGFeRD e-invoice validation.</summary>
    Mustang,

    /// <summary>Caller-supplied validator not otherwise modeled by OfficeIMO.Pdf.</summary>
    Custom
}
