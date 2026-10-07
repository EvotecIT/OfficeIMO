namespace OfficeIMO.Html;

/// <summary>
/// Selects the bounded user-agent defaults applied before authored CSS.
/// </summary>
public enum HtmlRenderUserAgentStyleMode {
    /// <summary>
    /// Preserves OfficeIMO's document-oriented defaults, including an uninset body and Arial fallback.
    /// </summary>
    Document = 0,

    /// <summary>
    /// Applies selected browser defaults, including an eight CSS-pixel body margin.
    /// This does not claim complete browser user-agent stylesheet compatibility.
    /// </summary>
    Browser = 1
}
