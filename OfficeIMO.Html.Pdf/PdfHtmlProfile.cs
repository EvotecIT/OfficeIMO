namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Selects how PDF logical content is exported as HTML.
/// </summary>
public enum PdfHtmlProfile {
    /// <summary>
    /// PDF logical content is exported as semantic HTML elements such as headings, paragraphs, lists, tables, and metadata placeholders.
    /// </summary>
    Semantic,

    /// <summary>
    /// PDF logical content is exported as review HTML with one page wrapper per source page and absolutely positioned text/table/image/link/form hints.
    /// </summary>
    PositionedReview
}
