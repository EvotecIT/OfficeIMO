namespace OfficeIMO.Markdown.Html;

/// <summary>
/// Options controlling HTML to Markdown conversion.
/// </summary>
public sealed class HtmlToMarkdownOptions {
    private readonly HashSet<string> _appliedPluginIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _appliedFeaturePackIds = new(StringComparer.OrdinalIgnoreCase);

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
    /// Optional ordered custom HTML element block converters used before the base converter falls
    /// back to generic block handling.
    /// </summary>
    public List<HtmlElementBlockConverter> ElementBlockConverters { get; } = new();

    /// <summary>
    /// Optional ordered custom HTML inline element converters used before the base converter falls
    /// back to generic inline handling.
    /// </summary>
    public List<HtmlInlineElementConverter> InlineElementConverters { get; } = new();

    /// <summary>
    /// Optional ordered visual-host round-trip hints used when shared <c>data-omd-*</c> elements
    /// should recover richer semantic fenced blocks than the default visual-contract mapping.
    /// </summary>
    public List<MarkdownVisualElementRoundTripHint> VisualElementRoundTripHints { get; } = new();

    /// <summary>
    /// Marks a renderer plugin's HTML-ingestion contract as applied to these options.
    /// </summary>
    public bool TryMarkPluginApplied(string pluginId) {
        if (string.IsNullOrWhiteSpace(pluginId)) {
            throw new ArgumentException("Plugin id is required.", nameof(pluginId));
        }

        return _appliedPluginIds.Add(pluginId.Trim());
    }

    /// <summary>
    /// Returns <see langword="true"/> when the supplied renderer plugin id has already been
    /// applied as an HTML-ingestion contract to these options.
    /// </summary>
    public bool HasPluginId(string pluginId) {
        if (string.IsNullOrWhiteSpace(pluginId)) {
            throw new ArgumentException("Plugin id is required.", nameof(pluginId));
        }

        return _appliedPluginIds.Contains(pluginId.Trim());
    }

    /// <summary>
    /// Marks a renderer feature pack's HTML-ingestion contract as applied to these options.
    /// </summary>
    public bool TryMarkFeaturePackApplied(string featurePackId) {
        if (string.IsNullOrWhiteSpace(featurePackId)) {
            throw new ArgumentException("Feature pack id is required.", nameof(featurePackId));
        }

        return _appliedFeaturePackIds.Add(featurePackId.Trim());
    }

    /// <summary>
    /// Returns <see langword="true"/> when the supplied renderer feature-pack id has already been
    /// applied as an HTML-ingestion contract to these options.
    /// </summary>
    public bool HasFeaturePackId(string featurePackId) {
        if (string.IsNullOrWhiteSpace(featurePackId)) {
            throw new ArgumentException("Feature pack id is required.", nameof(featurePackId));
        }

        return _appliedFeaturePackIds.Contains(featurePackId.Trim());
    }

    /// <summary>
    /// Compatibility wrapper for older visual round-trip tracking.
    /// </summary>
    public bool TryMarkVisualRoundTripPluginApplied(string pluginId) => TryMarkPluginApplied(pluginId);

    /// <summary>
    /// Compatibility wrapper for older visual round-trip tracking.
    /// </summary>
    public bool HasVisualRoundTripPluginId(string pluginId) => HasPluginId(pluginId);

    /// <summary>
    /// Compatibility wrapper for older visual round-trip tracking.
    /// </summary>
    public bool TryMarkVisualRoundTripFeaturePackApplied(string featurePackId) => TryMarkFeaturePackApplied(featurePackId);

    /// <summary>
    /// Compatibility wrapper for older visual round-trip tracking.
    /// </summary>
    public bool HasVisualRoundTripFeaturePackId(string featurePackId) => HasFeaturePackId(featurePackId);

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

        for (var i = 0; i < ElementBlockConverters.Count; i++) {
            var converter = ElementBlockConverters[i];
            if (converter != null) {
                clone.ElementBlockConverters.Add(converter);
            }
        }

        for (var i = 0; i < InlineElementConverters.Count; i++) {
            var converter = InlineElementConverters[i];
            if (converter != null) {
                clone.InlineElementConverters.Add(converter);
            }
        }

        for (var i = 0; i < VisualElementRoundTripHints.Count; i++) {
            var hint = VisualElementRoundTripHints[i];
            if (hint != null) {
                clone.VisualElementRoundTripHints.Add(hint);
            }
        }

        foreach (var pluginId in _appliedPluginIds) {
            clone._appliedPluginIds.Add(pluginId);
        }

        foreach (var featurePackId in _appliedFeaturePackIds) {
            clone._appliedFeaturePackIds.Add(featurePackId);
        }

        return clone;
    }
}
