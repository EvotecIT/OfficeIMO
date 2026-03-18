namespace OfficeIMO.Markdown;

/// <summary>
/// HTML-only responsive picture source metadata preserved on imported image blocks.
/// </summary>
public sealed class ImagePictureSource {
    /// <summary>Resolved source URL.</summary>
    public string Path { get; }

    /// <summary>Optional fully resolved srcset value preserved from the source element.</summary>
    public string? SrcSet { get; }

    /// <summary>Optional media query associated with the source.</summary>
    public string? Media { get; }

    /// <summary>Optional MIME type associated with the source.</summary>
    public string? Type { get; }

    /// <summary>Optional sizes hint associated with the source.</summary>
    public string? Sizes { get; }

    /// <summary>Create a responsive picture source entry.</summary>
    public ImagePictureSource(string path, string? media = null, string? type = null, string? sizes = null, string? srcSet = null) {
        Path = path ?? string.Empty;
        SrcSet = srcSet;
        Media = media;
        Type = type;
        Sizes = sizes;
    }
}
