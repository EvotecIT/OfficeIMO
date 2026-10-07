namespace OfficeIMO.Markup.Word;

/// <summary>Controls conversion of an Office Markup document to a Word document.</summary>
public sealed class MarkupToWordOptions {
    /// <summary>Base directory used to resolve relative resource paths.</summary>
    public string? BaseDirectory { get; set; }
    /// <summary>Whether image paths outside <see cref="BaseDirectory"/> may be read.</summary>
    public bool AllowExternalImagePaths { get; set; }
    /// <summary>Whether base64 PNG and JPEG data URI images may be embedded.</summary>
    public bool AllowDataUriImages { get; set; } = true;
    /// <summary>Maximum decoded data URI image size, in bytes.</summary>
    public long MaximumDataUriImageBytes { get; set; } = 16L * 1024 * 1024;
    /// <summary>Whether unsupported blocks should be preserved as visible text.</summary>
    public bool IncludeUnsupportedBlocksAsText { get; set; } = true;
    /// <summary>Default rendered chart width in pixels.</summary>
    public int DefaultChartWidthPixels { get; set; } = 640;
    /// <summary>Default rendered chart height in pixels.</summary>
    public int DefaultChartHeightPixels { get; set; } = 360;
}
