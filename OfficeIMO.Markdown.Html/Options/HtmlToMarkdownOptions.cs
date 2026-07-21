namespace OfficeIMO.Markdown.Html;

using AngleSharp.Dom;
using OfficeIMO.Html;

/// <summary>
/// Options controlling HTML to Markdown conversion.
/// </summary>
public sealed class HtmlToMarkdownOptions {
    private int _maxTableExpandedColumns = TableBlock.MaxEffectiveColumnCount;
    private readonly HashSet<string> _appliedPluginIds = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _appliedFeaturePackIds = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Creates the default OfficeIMO-flavored conversion profile.</summary>
    public static HtmlToMarkdownOptions CreateOfficeIMOProfile() => new HtmlToMarkdownOptions();

    /// <summary>Creates conversion options for the requested Markdown output profile.</summary>
    public static HtmlToMarkdownOptions CreateProfile(MarkdownOutputProfile profile) =>
        profile switch {
            MarkdownOutputProfile.OfficeIMO => CreateOfficeIMOProfile(),
            MarkdownOutputProfile.CommonMark => CreateCommonMarkProfile(),
            MarkdownOutputProfile.GitHubFlavoredMarkdown => CreateGitHubFlavoredMarkdownProfile(),
            MarkdownOutputProfile.Portable => CreatePortableProfile(),
            _ => throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unknown HTML-to-Markdown output profile.")
        };

    /// <summary>
    /// Creates a CommonMark-oriented conversion profile. The intermediate OfficeIMO document model is preserved,
    /// while markdown serialization avoids GitHub-only output where practical.
    /// </summary>
    public static HtmlToMarkdownOptions CreateCommonMarkProfile() => new HtmlToMarkdownOptions {
        EscapeMarkdownLineStarts = true,
        MarkdownWriteOptions = MarkdownWriteOptions.CreateCommonMarkProfile()
    };

    /// <summary>
    /// Creates a GitHub Flavored Markdown-oriented conversion profile for README and GitHub documentation output.
    /// </summary>
    public static HtmlToMarkdownOptions CreateGitHubFlavoredMarkdownProfile() => new HtmlToMarkdownOptions {
        SmartHref = true,
        MarkdownWriteOptions = MarkdownWriteOptions.CreateGitHubFlavoredMarkdownProfile()
    };

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
    /// When true, links whose label already represents the target are emitted as plain text.
    /// </summary>
    public bool SmartHref { get; set; }

    /// <summary>
    /// Controls how unsupported block-level elements are converted.
    /// </summary>
    public HtmlUnknownTagHandling UnknownBlockHandling { get; set; } = HtmlUnknownTagHandling.Preserve;

    /// <summary>
    /// Controls how unsupported inline elements are converted.
    /// </summary>
    public HtmlUnknownTagHandling UnknownInlineHandling { get; set; } = HtmlUnknownTagHandling.Preserve;

    /// <summary>
    /// When true, plain text imported from HTML is escaped when it could be interpreted as a Markdown block marker.
    /// </summary>
    public bool EscapeMarkdownLineStarts { get; set; }

    /// <summary>
    /// Shared URL policy applied before hyperlink URLs are materialized into the Markdown model.
    /// </summary>
    public HtmlUrlPolicy UrlPolicy { get; set; } = HtmlUrlPolicy.CreateOfficeIMOProfile();

    /// <summary>
    /// Optional separate policy for image and media sources. When omitted, <see cref="UrlPolicy"/>
    /// is used so existing callers retain their single-policy behavior.
    /// </summary>
    public HtmlUrlPolicy? ResourceUrlPolicy { get; set; }

    /// <summary>
    /// Controls how base64 data-URI images are represented in the converted image model.
    /// </summary>
    public HtmlBase64ImageHandling Base64Images { get; set; } = HtmlBase64ImageHandling.Include;

    /// <summary>
    /// Output directory used when <see cref="Base64Images"/> is <see cref="HtmlBase64ImageHandling.SaveToFile"/>.
    /// </summary>
    public string? Base64ImageOutputDirectory { get; set; }

    /// <summary>
    /// Optional file name factory used when saving base64 image data. Arguments are the zero-based image index and MIME type.
    /// </summary>
    public Func<int, string, string>? Base64ImageFileNameGenerator { get; set; }

    /// <summary>
    /// Controls whether low-value metadata inside repeated listing cards should be preserved or suppressed.
    /// </summary>
    public HtmlListingCardMetadataMode ListingCardMetadataMode { get; set; } = HtmlListingCardMetadataMode.Preserve;

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
    /// Maximum logical columns produced by expanding HTML table colspan attributes.
    /// Values must be between 1 and 4096.
    /// </summary>
    public int MaxTableExpandedColumns {
        get => _maxTableExpandedColumns;
        set {
            if (value < 1 || value > TableBlock.MaxEffectiveColumnCount) {
                throw new ArgumentOutOfRangeException(nameof(MaxTableExpandedColumns), value, "MaxTableExpandedColumns must be between 1 and 4096.");
            }

            _maxTableExpandedColumns = value;
        }
    }

    /// <summary>
    /// Optional ordered post-conversion document transforms applied to the intermediate <see cref="MarkdownDoc"/>.
    /// </summary>
    /// <example>
    /// <code>
    /// var options = HtmlToMarkdownOptions.CreatePortableProfile();
    /// options.DocumentTransforms.Add(
    ///     new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence));
    ///
    /// var document = HtmlConversionDocument.Parse(html).ToMarkdownDocument(options);
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
    /// CSS selectors for elements that should be removed before conversion.
    /// </summary>
    public HashSet<string> ExcludeSelectors { get; } = new(StringComparer.Ordinal);

    /// <summary>
    /// Predicates for elements that should be removed before conversion. Return <see langword="true"/> to remove an element.
    /// </summary>
    public List<Func<IElement, bool>> ElementFilters { get; } = new();

    /// <summary>
    /// Maps an HTML tag name to another tag name before built-in conversion is selected.
    /// </summary>
    public Dictionary<string, string> TagAliases { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Tag names that should be emitted as raw HTML without converting their children.
    /// </summary>
    public HashSet<string> PassThroughTags { get; } = new(StringComparer.OrdinalIgnoreCase);

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
            SmartHref = SmartHref,
            UnknownBlockHandling = UnknownBlockHandling,
            UnknownInlineHandling = UnknownInlineHandling,
            EscapeMarkdownLineStarts = EscapeMarkdownLineStarts,
            UrlPolicy = UrlPolicy?.Clone() ?? HtmlUrlPolicy.CreateOfficeIMOProfile(),
            ResourceUrlPolicy = ResourceUrlPolicy?.Clone(),
            Base64Images = Base64Images,
            Base64ImageOutputDirectory = Base64ImageOutputDirectory,
            Base64ImageFileNameGenerator = Base64ImageFileNameGenerator,
            ListingCardMetadataMode = ListingCardMetadataMode,
            MarkdownWriteOptions = MarkdownWriteOptions?.Clone(),
            MaxInputCharacters = MaxInputCharacters,
            MaxTableExpandedColumns = MaxTableExpandedColumns
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

        foreach (var selector in ExcludeSelectors) {
            if (!string.IsNullOrWhiteSpace(selector)) {
                clone.ExcludeSelectors.Add(selector);
            }
        }

        for (var i = 0; i < ElementFilters.Count; i++) {
            var filter = ElementFilters[i];
            if (filter != null) {
                clone.ElementFilters.Add(filter);
            }
        }

        foreach (var alias in TagAliases) {
            if (!string.IsNullOrWhiteSpace(alias.Key) && !string.IsNullOrWhiteSpace(alias.Value)) {
                clone.TagAliases[alias.Key] = alias.Value;
            }
        }

        foreach (var tagName in PassThroughTags) {
            if (!string.IsNullOrWhiteSpace(tagName)) {
                clone.PassThroughTags.Add(tagName);
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
