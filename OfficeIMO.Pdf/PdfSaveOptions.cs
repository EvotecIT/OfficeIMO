namespace OfficeIMO.Pdf;

/// <summary>
/// Provides configuration options when saving a document to PDF.
/// </summary>
public class PdfSaveOptions {
    /// <summary>
    /// Left page margin in centimeters. Defaults to 1 cm.
    /// </summary>
    public float MarginLeft { get; set; } = 1f;

    /// <summary>
    /// Top page margin in centimeters. Defaults to 1 cm.
    /// </summary>
    public float MarginTop { get; set; } = 1f;

    /// <summary>
    /// Right page margin in centimeters. Defaults to 1 cm.
    /// </summary>
    public float MarginRight { get; set; } = 1f;

    /// <summary>
    /// Bottom page margin in centimeters. Defaults to 1 cm.
    /// </summary>
    public float MarginBottom { get; set; } = 1f;
}
