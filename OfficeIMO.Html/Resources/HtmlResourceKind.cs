namespace OfficeIMO.Html;

/// <summary>
/// Resource dependency type discovered in HTML input.
/// </summary>
public enum HtmlResourceKind {
    /// <summary>Image resource.</summary>
    Image,
    /// <summary>Stylesheet resource.</summary>
    Stylesheet,
    /// <summary>Hyperlink target.</summary>
    Hyperlink,
    /// <summary>Script resource.</summary>
    Script,
    /// <summary>Audio, video, source, or track resource.</summary>
    Media,
    /// <summary>Font or preloaded font-like resource.</summary>
    Font,
    /// <summary>Resource that does not fit a more specific category.</summary>
    Other
}
