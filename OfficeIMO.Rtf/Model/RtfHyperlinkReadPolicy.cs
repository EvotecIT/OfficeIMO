namespace OfficeIMO.Rtf;

/// <summary>Controls which hyperlink field targets are materialized during semantic RTF binding.</summary>
public enum RtfHyperlinkReadPolicy {
    /// <summary>Preserves every hyperlink target without fetching it.</summary>
    AllowAll,

    /// <summary>Allows relative links, fragments, HTTP, HTTPS, and mail links.</summary>
    WebAndMailOnly,

    /// <summary>Flattens every hyperlink field to its visible result.</summary>
    BlockAll
}
