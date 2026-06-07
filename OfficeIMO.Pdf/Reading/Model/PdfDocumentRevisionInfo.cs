namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight cross-reference revision marker read from a PDF file.
/// </summary>
public sealed class PdfDocumentRevisionInfo {
    internal PdfDocumentRevisionInfo(int revisionNumber, int startXrefOffset, int? previousXrefOffset) {
        RevisionNumber = revisionNumber;
        StartXrefOffset = startXrefOffset;
        PreviousXrefOffset = previousXrefOffset;
    }

    /// <summary>One-based revision number in file order.</summary>
    public int RevisionNumber { get; }

    /// <summary>Offset declared by this revision's startxref section.</summary>
    public int StartXrefOffset { get; }

    /// <summary>Previous xref offset linked from this revision's trailer or xref stream, when readable.</summary>
    public int? PreviousXrefOffset { get; }

    /// <summary>True when this revision links to a previous revision.</summary>
    public bool HasPreviousRevision => PreviousXrefOffset.HasValue;
}
