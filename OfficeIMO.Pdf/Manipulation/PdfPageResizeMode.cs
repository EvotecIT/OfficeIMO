namespace OfficeIMO.Pdf;

/// <summary>
/// Controls how existing page content is scaled when a PDF page is resized.
/// </summary>
public enum PdfPageResizeMode {
    /// <summary>Preserve the original aspect ratio and fit the full source page inside the target page.</summary>
    Fit,

    /// <summary>Preserve the original aspect ratio and fill the target page, allowing content to be clipped.</summary>
    Fill,

    /// <summary>Scale width and height independently so the source page exactly fills the target page.</summary>
    Stretch
}
