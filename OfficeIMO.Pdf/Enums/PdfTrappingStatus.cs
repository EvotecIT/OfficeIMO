namespace OfficeIMO.Pdf;

/// <summary>
/// Print trapping status written to the PDF document information dictionary.
/// </summary>
public enum PdfTrappingStatus {
    /// <summary>The document trapping state is not known.</summary>
    Unknown,
    /// <summary>The document has not been trapped.</summary>
    False,
    /// <summary>The document has been trapped.</summary>
    True
}
