namespace OfficeIMO.Pdf;

/// <summary>
/// Standard PDF page boundary box names used by viewer preferences.
/// </summary>
public enum PdfPageBoundaryBox {
    /// <summary>
    /// Uses the page media box.
    /// </summary>
    MediaBox,

    /// <summary>
    /// Uses the page crop box.
    /// </summary>
    CropBox,

    /// <summary>
    /// Uses the page bleed box.
    /// </summary>
    BleedBox,

    /// <summary>
    /// Uses the page trim box.
    /// </summary>
    TrimBox,

    /// <summary>
    /// Uses the page art box.
    /// </summary>
    ArtBox
}
