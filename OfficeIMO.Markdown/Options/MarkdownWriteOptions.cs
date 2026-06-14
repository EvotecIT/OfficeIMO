namespace OfficeIMO.Markdown;

/// <summary>
/// Options controlling markdown serialization from the typed document model.
/// </summary>
public sealed class MarkdownWriteOptions {
    private string? _outputLineEnding;
    private char _unorderedListMarker = '-';

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
    /// Optional line ending to use in rendered Markdown. When unset, the platform default is used.
    /// </summary>
    public string? OutputLineEnding {
        get => _outputLineEnding;
        set {
            if (value != null && value.Length == 0) {
                throw new ArgumentException("Output line ending cannot be empty.", nameof(value));
            }

            _outputLineEnding = value;
        }
    }

    /// <summary>
    /// Marker used when rendering unordered list items. CommonMark-compatible values are <c>-</c>, <c>*</c>, and <c>+</c>.
    /// </summary>
    public char UnorderedListMarker {
        get => _unorderedListMarker;
        set {
            if (value != '-' && value != '*' && value != '+') {
                throw new ArgumentOutOfRangeException(nameof(value), value, "Unordered list marker must be '-', '*', or '+'.");
            }

            _unorderedListMarker = value;
        }
    }

    /// <summary>
    /// Optional markdown block render extensions. Later registrations win when block types overlap.
    /// </summary>
    public List<MarkdownBlockMarkdownRenderExtension> BlockRenderExtensions { get; } = new();

    /// <summary>
    /// Creates a shallow clone of the writer options while copying mutable collections.
    /// </summary>
    public MarkdownWriteOptions Clone() {
        var clone = new MarkdownWriteOptions {
            ImageRenderingMode = ImageRenderingMode,
            OutputLineEnding = OutputLineEnding,
            UnorderedListMarker = UnorderedListMarker
        };
        for (int i = 0; i < BlockRenderExtensions.Count; i++) {
            clone.BlockRenderExtensions.Add(BlockRenderExtensions[i]);
        }

        return clone;
    }
}
