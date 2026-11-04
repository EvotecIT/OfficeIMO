namespace OfficeIMO.Html;

/// <summary>
/// Selects the bounded user-agent defaults applied before authored CSS.
/// </summary>
public enum HtmlRenderUserAgentStyleMode {
    /// <summary>
    /// Preserves OfficeIMO's document-oriented defaults, including Arial fallback and, outside
    /// direct HTML-to-PDF rendering, an uninset body. <c>new HtmlPdfExportOptions()</c> from OfficeIMO.Html.Pdf keeps the browser's
    /// eight CSS-pixel body margin so PDF output wraps as it does in a browser; author CSS such as
    /// <c>body{margin:0}</c> overrides it.
    /// </summary>
    Document = 0,

    /// <summary>
    /// Applies selected browser defaults, including an eight CSS-pixel body margin.
    /// This does not claim complete browser user-agent stylesheet compatibility.
    /// </summary>
    Browser = 1
}
