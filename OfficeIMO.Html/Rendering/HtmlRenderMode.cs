namespace OfficeIMO.Html;

/// <summary>
/// Selects the layout surface used by the first-party HTML renderer.
/// </summary>
public enum HtmlRenderMode {
    /// <summary>Renders the document as one continuous screen-oriented surface.</summary>
    Continuous,

    /// <summary>Fragments the document into fixed print-oriented pages.</summary>
    Paged
}
