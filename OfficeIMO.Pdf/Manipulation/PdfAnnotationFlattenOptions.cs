namespace OfficeIMO.Pdf;

/// <summary>Selects visual annotations to paint into page content and remove as live annotations.</summary>
public sealed class PdfAnnotationFlattenOptions {
    /// <summary>Optional exact indirect annotation object number.</summary>
    public int? ObjectNumber { get; set; }
    /// <summary>Optional one-based page number.</summary>
    public int? PageNumber { get; set; }
    /// <summary>Optional annotation subtype.</summary>
    public string? Subtype { get; set; }
}
