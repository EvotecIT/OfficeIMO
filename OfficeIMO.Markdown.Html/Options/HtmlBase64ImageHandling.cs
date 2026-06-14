namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Controls how HTML data-URI images are handled during HTML-to-Markdown conversion.
/// </summary>
public enum HtmlBase64ImageHandling {
    /// <summary>Keep base64 data-URI image sources in the generated Markdown image model.</summary>
    Include = 0,

    /// <summary>Drop base64 data-URI images from the converted document.</summary>
    Skip = 1,

    /// <summary>Decode base64 data-URI images into files and use the saved file path in the Markdown image model.</summary>
    SaveToFile = 2
}
