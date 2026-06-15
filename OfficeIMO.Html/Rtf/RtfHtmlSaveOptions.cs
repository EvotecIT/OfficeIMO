namespace OfficeIMO.Html;

/// <summary>
/// Controls RTF to semantic HTML conversion.
/// </summary>
public sealed class RtfHtmlSaveOptions {
    /// <summary>Writes only the body fragment instead of a complete HTML document.</summary>
    public bool FragmentOnly { get; set; } = true;

    /// <summary>Includes document metadata when a full HTML document is requested.</summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>Optional HTML document title. When unset, the RTF title is used.</summary>
    public string? Title { get; set; }

    /// <summary>Embeds PNG and JPEG images as data URI values.</summary>
    public bool EmbedImagesAsDataUri { get; set; } = true;

    /// <summary>Newline sequence used by the generated HTML.</summary>
    public string NewLine { get; set; } = Environment.NewLine;

    /// <summary>
    /// Creates a reusable copy of the current save options.
    /// </summary>
    /// <returns>A new <see cref="RtfHtmlSaveOptions"/> with the same configuration values.</returns>
    public RtfHtmlSaveOptions Clone() => new RtfHtmlSaveOptions {
        FragmentOnly = FragmentOnly,
        IncludeMetadata = IncludeMetadata,
        Title = Title,
        EmbedImagesAsDataUri = EmbedImagesAsDataUri,
        NewLine = NewLine
    };

    internal string GetNewLine() => string.IsNullOrEmpty(NewLine) ? Environment.NewLine : NewLine;
}
