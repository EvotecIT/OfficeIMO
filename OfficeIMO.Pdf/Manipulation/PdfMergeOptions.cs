namespace OfficeIMO.Pdf;

/// <summary>
/// Configures optional behavior for first-party PDF merge operations.
/// </summary>
public sealed class PdfMergeOptions {
    /// <summary>
    /// Gets or sets whether supported visual annotations should be painted into page content before pages are merged.
    /// Link annotations, form fields, and unsupported annotation shapes remain unchanged unless another OfficeIMO.Pdf operation handles them.
    /// </summary>
    public bool FlattenVisualAnnotations { get; set; }

    /// <summary>
    /// Gets or sets page resize options applied to each input PDF before pages are merged.
    /// Use this to normalize mixed source page sizes into a fixed output paper size.
    /// </summary>
    public PdfPageResizeOptions? ResizePages { get; set; }
}
