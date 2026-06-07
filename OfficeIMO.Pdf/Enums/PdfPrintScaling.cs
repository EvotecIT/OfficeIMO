namespace OfficeIMO.Pdf;

/// <summary>
/// Viewer print-scaling preference for generated PDFs.
/// </summary>
public enum PdfPrintScaling {
    /// <summary>
    /// Lets the viewer apply its default print scaling.
    /// </summary>
    AppDefault,

    /// <summary>
    /// Requests no viewer-side print scaling.
    /// </summary>
    None
}
