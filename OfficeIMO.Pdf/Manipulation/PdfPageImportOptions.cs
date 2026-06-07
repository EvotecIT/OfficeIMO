namespace OfficeIMO.Pdf;

/// <summary>
/// Configures optional behavior for first-party PDF page import operations.
/// </summary>
public sealed class PdfPageImportOptions {
    /// <summary>
    /// Gets or sets whether supported visual annotations from imported source pages should be painted into page content.
    /// Existing target-document annotations are not flattened by this option.
    /// </summary>
    public bool FlattenVisualAnnotations { get; set; }
}
