namespace OfficeIMO.Pdf;

/// <summary>
/// Relationship between a PDF document and an associated embedded file.
/// </summary>
public enum PdfAssociatedFileRelationship {
    /// <summary>The relationship is intentionally unspecified.</summary>
    Unspecified = 0,
    /// <summary>The embedded file is the source material for the PDF.</summary>
    Source,
    /// <summary>The embedded file contains data associated with the PDF.</summary>
    Data,
    /// <summary>The embedded file is an alternative representation of the PDF content.</summary>
    Alternative,
    /// <summary>The embedded file supplements the PDF content.</summary>
    Supplement
}
