namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Controls how HTML-to-Markdown conversion handles elements that do not map to a built-in Markdown construct.
/// </summary>
public enum HtmlUnknownTagHandling {
    /// <summary>Preserve the original HTML for the unsupported element when the matching preserve option is enabled.</summary>
    Preserve,

    /// <summary>Drop the unsupported element and its children.</summary>
    Drop,

    /// <summary>Ignore the unsupported element wrapper and convert its children.</summary>
    Bypass,

    /// <summary>Fail conversion when an unsupported element is encountered.</summary>
    Raise
}
