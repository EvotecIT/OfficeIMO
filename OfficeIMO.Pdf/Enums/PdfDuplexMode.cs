namespace OfficeIMO.Pdf;

/// <summary>
/// Viewer duplex-printing preference for generated PDFs.
/// </summary>
public enum PdfDuplexMode {
    /// <summary>
    /// Requests one-sided printing.
    /// </summary>
    Simplex,

    /// <summary>
    /// Requests two-sided printing flipped on the short edge.
    /// </summary>
    DuplexFlipShortEdge,

    /// <summary>
    /// Requests two-sided printing flipped on the long edge.
    /// </summary>
    DuplexFlipLongEdge
}
