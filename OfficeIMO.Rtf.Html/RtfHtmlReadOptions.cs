namespace OfficeIMO.Rtf.Html;

/// <summary>
/// Controls semantic HTML to RTF conversion.
/// </summary>
public sealed class RtfHtmlReadOptions {
    /// <summary>Base URI used to resolve relative hyperlinks and image sources.</summary>
    public Uri? BaseUri { get; set; }

    /// <summary>Preserves unknown element names as bracketed text markers instead of treating them as transparent containers.</summary>
    public bool PreserveUnknownTagsAsText { get; set; }

    /// <summary>When enabled, text nodes made only of whitespace are ignored outside preformatted elements.</summary>
    public bool IgnoreInsignificantWhitespace { get; set; } = true;
}
