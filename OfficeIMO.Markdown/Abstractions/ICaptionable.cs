namespace OfficeIMO.Markdown;

/// <summary>
/// Internal capability for blocks that can display a caption (e.g., images, code blocks).
/// </summary>
internal interface ICaptionable {
    /// <summary>Optional caption text.</summary>
    string? Caption { get; set; }
}
