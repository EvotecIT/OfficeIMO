namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Selects the internal first-party path used by the HTML to PDF adapter.
/// </summary>
public enum HtmlPdfProfile {
    /// <summary>
    /// HTML is converted into the OfficeIMO Markdown AST, then rendered through the Markdown PDF adapter.
    /// </summary>
    Semantic,

    /// <summary>
    /// HTML is converted into an OfficeIMO Word document, then rendered through the Word PDF adapter.
    /// </summary>
    Document,

    /// <summary>
    /// HTML is laid out directly by OfficeIMO.Html and projected into OfficeIMO.Pdf without a Word or Markdown intermediate.
    /// </summary>
    Rendered
}
