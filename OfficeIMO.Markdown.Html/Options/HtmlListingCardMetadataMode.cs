namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Controls how low-value metadata inside repeated listing cards is treated during HTML-to-markdown conversion.
/// </summary>
public enum HtmlListingCardMetadataMode {
    /// <summary>
    /// Preserves all listing-card metadata elements.
    /// </summary>
    Preserve = 0,

    /// <summary>
    /// Suppresses low-value metadata like date/time/byline/read-more blocks when they appear inside repeated listing cards.
    /// </summary>
    SuppressInRepeatedCards = 1
}
