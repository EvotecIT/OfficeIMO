namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Controls how PDF images are represented in exported HTML.
/// </summary>
public enum PdfHtmlImageExportMode {
    /// <summary>
    /// Emit readable placeholders only.
    /// </summary>
    PlaceholderOnly = 0,

    /// <summary>
    /// Embed extracted image files as data URI images when available, falling back to placeholders otherwise.
    /// </summary>
    EmbeddedDataUri = 1
}
