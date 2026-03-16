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
        MarkdownBlockRenderBuiltInExtensions.AddPortableCalloutMarkdownFallback(options);
        return options;
    }

    /// <summary>
    /// Optional markdown block render extensions. Later registrations win when block types overlap.
    /// </summary>
    public List<MarkdownBlockMarkdownRenderExtension> BlockRenderExtensions { get; } = new();

    /// <summary>
    /// Creates a shallow clone of the writer options while copying mutable collections.
    /// </summary>
    public MarkdownWriteOptions Clone() {
        var clone = new MarkdownWriteOptions();
        for (int i = 0; i < BlockRenderExtensions.Count; i++) {
            if (BlockRenderExtensions[i] != null) {
                clone.BlockRenderExtensions.Add(BlockRenderExtensions[i]);
            }
        }

        return clone;
    }
}
