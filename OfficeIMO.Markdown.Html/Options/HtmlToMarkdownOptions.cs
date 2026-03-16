namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Options controlling HTML to Markdown conversion.
/// </summary>
public sealed class HtmlToMarkdownOptions {
    /// <summary>Creates the default OfficeIMO-flavored conversion profile.</summary>
    public static HtmlToMarkdownOptions CreateOfficeIMOProfile() => new HtmlToMarkdownOptions();

    /// <summary>
    /// Creates a portable conversion profile that serializes the converted document with portable markdown fallbacks.
    /// </summary>
    public static HtmlToMarkdownOptions CreatePortableProfile() => new HtmlToMarkdownOptions {
        MarkdownWriteOptions = MarkdownWriteOptions.CreatePortableProfile()
    };

    /// <summary>
    /// Optional base URI used to resolve relative links and image sources.
    /// </summary>
    public Uri? BaseUri { get; set; }

    /// <summary>
    /// When true, only the body contents are converted when a body element is present.
    /// </summary>
    public bool UseBodyContentsOnly { get; set; } = true;

    /// <summary>
    /// When true, script/style/noscript/template elements are ignored.
    /// </summary>
    public bool RemoveScriptsAndStyles { get; set; } = true;

    /// <summary>
    /// When true, unsupported block elements are emitted as raw HTML blocks instead of being dropped.
    /// </summary>
    public bool PreserveUnsupportedBlocks { get; set; } = true;

    /// <summary>
    /// When true, unsupported inline elements are emitted as raw HTML inside inline Markdown.
    /// </summary>
    public bool PreserveUnsupportedInlineHtml { get; set; } = true;

    /// <summary>
    /// Optional markdown writer options used when the converter serializes the intermediate
    /// <see cref="MarkdownDoc"/> back to markdown text.
    /// </summary>
    public MarkdownWriteOptions? MarkdownWriteOptions { get; set; }

    /// <summary>
    /// Optional maximum input length, in characters, accepted by HTML-to-markdown conversion.
    /// When set and exceeded, conversion fails fast with an <see cref="ArgumentOutOfRangeException"/>.
    /// </summary>
    public int? MaxInputCharacters { get; set; }

    /// <summary>
    /// Optional ordered post-conversion document transforms applied to the intermediate <see cref="MarkdownDoc"/>.
    /// </summary>
    /// <example>
    /// <code>
    /// var options = HtmlToMarkdownOptions.CreatePortableProfile();
    /// options.DocumentTransforms.Add(
    ///     new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));
    ///
    /// var document = html.LoadFromHtml(options);
    /// </code>
    /// </example>
    public List<IMarkdownDocumentTransform> DocumentTransforms { get; } = new();

    /// <summary>
    /// Creates a copy of the current options instance so callers can reuse option templates safely.
    /// </summary>
    /// <returns>A new <see cref="HtmlToMarkdownOptions"/> with the same option values.</returns>
    public HtmlToMarkdownOptions Clone() {
        var clone = new HtmlToMarkdownOptions {
            BaseUri = BaseUri,
            UseBodyContentsOnly = UseBodyContentsOnly,
            RemoveScriptsAndStyles = RemoveScriptsAndStyles,
            PreserveUnsupportedBlocks = PreserveUnsupportedBlocks,
            PreserveUnsupportedInlineHtml = PreserveUnsupportedInlineHtml,
            MarkdownWriteOptions = MarkdownWriteOptions?.Clone(),
            MaxInputCharacters = MaxInputCharacters
        };

        for (var i = 0; i < DocumentTransforms.Count; i++) {
            var transform = DocumentTransforms[i];
            if (transform != null) {
                clone.DocumentTransforms.Add(transform);
            }
        }

        return clone;
    }
}
