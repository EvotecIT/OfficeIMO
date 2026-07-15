namespace OfficeIMO.Pdf;

/// <summary>Safe, immutable, bounded view of the active PDF object graph and revision metadata.</summary>
public sealed class PdfRawDocumentView {
    internal PdfRawDocumentView(
        IReadOnlyList<PdfRawObjectView> objects,
        int totalObjectCount,
        int? catalogObjectNumber,
        string trailerPreview,
        bool isTruncated,
        IReadOnlyList<PdfDocumentRevisionInfo> revisions) {
        Objects = objects;
        TotalObjectCount = totalObjectCount;
        CatalogObjectNumber = catalogObjectNumber;
        TrailerPreview = trailerPreview;
        IsTruncated = isTruncated;
        Revisions = revisions;
    }

    /// <summary>Projected active indirect objects ordered by object number.</summary>
    public IReadOnlyList<PdfRawObjectView> Objects { get; }
    /// <summary>Total parsed active object count before projection limits.</summary>
    public int TotalObjectCount { get; }
    /// <summary>Catalog object number when it can be resolved.</summary>
    public int? CatalogObjectNumber { get; }
    /// <summary>Bounded active trailer-chain syntax preview.</summary>
    public string TrailerPreview { get; }
    /// <summary>True when object or trailer projection was truncated.</summary>
    public bool IsTruncated { get; }
    /// <summary>Read-only revision-chain metadata discovered by the parser.</summary>
    public IReadOnlyList<PdfDocumentRevisionInfo> Revisions { get; }

    /// <summary>Returns an active indirect object by number, or null when not projected.</summary>
    public PdfRawObjectView? GetObject(int objectNumber) {
        Guard.PositiveInteger(objectNumber, nameof(objectNumber));
        for (int i = 0; i < Objects.Count; i++) {
            if (Objects[i].ObjectNumber == objectNumber) return Objects[i];
        }

        return null;
    }
}
