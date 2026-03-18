namespace OfficeIMO.Markdown;

/// <summary>
/// Options controlling markdown serialization from the typed document model.
/// </summary>
public sealed class MarkdownWriteOptions {
    /// <summary>Creates default OfficeIMO-flavored markdown writer options.</summary>
    public static MarkdownWriteOptions CreateOfficeIMOProfile() => new MarkdownWriteOptions();

    /// <summary>
    /// Creates a more portable markdown writer profile that degrades OfficeIMO-only blocks into broadly compatible markdown.
    /// </summary>
    public static MarkdownWriteOptions CreatePortableProfile() {
        var options = new MarkdownWriteOptions();
        options.ImageRenderingMode = MarkdownImageRenderingMode.PortableMarkdown;
        MarkdownBlockRenderBuiltInExtensions.AddPortableCalloutMarkdownFallback(options);
        return options;
    }

    /// <summary>
    /// Creates a markdown writer profile that emits raw HTML for image output.
    /// </summary>
    public static MarkdownWriteOptions CreateHtmlImageProfile() => new MarkdownWriteOptions {
        ImageRenderingMode = MarkdownImageRenderingMode.Html
    };

    /// <summary>
    /// Controls how image blocks and image inlines are serialized back to markdown text.
    /// </summary>
    public MarkdownImageRenderingMode ImageRenderingMode { get; set; } = MarkdownImageRenderingMode.RichMarkdown;

    /// <summary>
    /// Optional markdown block render extensions. Later registrations win when block types overlap.
    /// </summary>
    public List<MarkdownBlockMarkdownRenderExtension> BlockRenderExtensions { get; } = new();

    /// <summary>
    /// Creates a shallow clone of the writer options while copying mutable collections.
    /// </summary>
    public MarkdownWriteOptions Clone() {
        var clone = new MarkdownWriteOptions {
            ImageRenderingMode = ImageRenderingMode
        };
        for (int i = 0; i < BlockRenderExtensions.Count; i++) {
            if (BlockRenderExtensions[i] != null) {
                clone.BlockRenderExtensions.Add(BlockRenderExtensions[i]);
            }
        }

        return clone;
    }
}
