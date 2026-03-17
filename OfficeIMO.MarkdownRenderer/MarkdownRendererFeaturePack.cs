using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Reusable host-level feature pack that can apply a coordinated set of renderer plugins,
/// preprocessors, postprocessors, and option defaults in one idempotent step.
/// </summary>
public sealed class MarkdownRendererFeaturePack {
    private readonly Action<MarkdownRendererOptions> _apply;
    private readonly Action<MarkdownReaderOptions> _applyReader;
    private readonly IReadOnlyList<MarkdownRendererPlugin> _plugins;
    private readonly IReadOnlyList<IMarkdownDocumentTransform> _readerDocumentTransforms;
    private readonly IReadOnlyList<IMarkdownDocumentTransform> _htmlDocumentTransforms;
    private readonly IReadOnlyList<IMarkdownDocumentTransform> _rendererDocumentTransforms;
    private readonly IReadOnlyList<HtmlElementBlockConverter> _htmlElementBlockConverters;
    private readonly IReadOnlyList<HtmlInlineElementConverter> _htmlInlineElementConverters;
    private readonly IReadOnlyList<MarkdownVisualElementRoundTripHint> _visualElementRoundTripHints;

    /// <summary>
    /// Creates a new host-level feature pack.
    /// </summary>
    public MarkdownRendererFeaturePack(
        string id,
        string name,
        Action<MarkdownRendererOptions> apply,
        Action<MarkdownReaderOptions>? applyReader = null) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Feature pack id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Feature pack name is required.", nameof(name));
        }

        _apply = apply ?? throw new ArgumentNullException(nameof(apply));
        _plugins = Array.Empty<MarkdownRendererPlugin>();
        _readerDocumentTransforms = Array.Empty<IMarkdownDocumentTransform>();
        _htmlDocumentTransforms = Array.Empty<IMarkdownDocumentTransform>();
        _rendererDocumentTransforms = Array.Empty<IMarkdownDocumentTransform>();
        _htmlElementBlockConverters = Array.Empty<HtmlElementBlockConverter>();
        _htmlInlineElementConverters = Array.Empty<HtmlInlineElementConverter>();
        _visualElementRoundTripHints = Array.Empty<MarkdownVisualElementRoundTripHint>();
        _applyReader = applyReader ?? (_ => { });
        Id = id.Trim();
        Name = name.Trim();
    }

    /// <summary>
    /// Creates a new host-level feature pack from reusable plugins plus an optional custom apply step.
    /// </summary>
    public MarkdownRendererFeaturePack(
        string id,
        string name,
        IEnumerable<MarkdownRendererPlugin> plugins,
        Action<MarkdownRendererOptions>? apply = null,
        Action<MarkdownReaderOptions>? applyReader = null) {
        if (string.IsNullOrWhiteSpace(id)) {
            throw new ArgumentException("Feature pack id is required.", nameof(id));
        }

        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Feature pack name is required.", nameof(name));
        }

        if (plugins == null) {
            throw new ArgumentNullException(nameof(plugins));
        }

        var normalizedPlugins = new List<MarkdownRendererPlugin>();
        var readerDocumentTransforms = new List<IMarkdownDocumentTransform>();
        var htmlDocumentTransforms = new List<IMarkdownDocumentTransform>();
        var rendererDocumentTransforms = new List<IMarkdownDocumentTransform>();
        var htmlElementBlockConverters = new List<HtmlElementBlockConverter>();
        var htmlInlineElementConverters = new List<HtmlInlineElementConverter>();
        var roundTripHints = new List<MarkdownVisualElementRoundTripHint>();
        foreach (var plugin in plugins) {
            if (plugin == null) {
                continue;
            }

            bool exists = false;
            for (int i = 0; i < normalizedPlugins.Count; i++) {
                if (string.Equals(normalizedPlugins[i].Name, plugin.Name, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                normalizedPlugins.Add(plugin);
                for (int i = 0; i < plugin.ReaderDocumentTransforms.Count; i++) {
                    var transform = plugin.ReaderDocumentTransforms[i];
                    if (transform == null) {
                        continue;
                    }

                    bool transformExists = false;
                    for (int j = 0; j < readerDocumentTransforms.Count; j++) {
                        if (ReferenceEquals(readerDocumentTransforms[j], transform)) {
                            transformExists = true;
                            break;
                        }
                    }

                    if (!transformExists) {
                        readerDocumentTransforms.Add(transform);
                    }
                }

                for (int i = 0; i < plugin.HtmlDocumentTransforms.Count; i++) {
                    var transform = plugin.HtmlDocumentTransforms[i];
                    if (transform == null) {
                        continue;
                    }

                    bool transformExists = false;
                    for (int j = 0; j < htmlDocumentTransforms.Count; j++) {
                        if (ReferenceEquals(htmlDocumentTransforms[j], transform)) {
                            transformExists = true;
                            break;
                        }
                    }

                    if (!transformExists) {
                        htmlDocumentTransforms.Add(transform);
                    }
                }

                for (int i = 0; i < plugin.RendererDocumentTransforms.Count; i++) {
                    var transform = plugin.RendererDocumentTransforms[i];
                    if (transform == null) {
                        continue;
                    }

                    bool transformExists = false;
                    for (int j = 0; j < rendererDocumentTransforms.Count; j++) {
                        if (ReferenceEquals(rendererDocumentTransforms[j], transform)) {
                            transformExists = true;
                            break;
                        }
                    }

                    if (!transformExists) {
                        rendererDocumentTransforms.Add(transform);
                    }
                }

                for (int i = 0; i < plugin.HtmlElementBlockConverters.Count; i++) {
                    var converter = plugin.HtmlElementBlockConverters[i];
                    if (converter == null) {
                        continue;
                    }

                    bool converterExists = false;
                    for (int j = 0; j < htmlElementBlockConverters.Count; j++) {
                        if (string.Equals(htmlElementBlockConverters[j].Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                            converterExists = true;
                            break;
                        }
                    }

                    if (!converterExists) {
                        htmlElementBlockConverters.Add(converter);
                    }
                }

                for (int i = 0; i < plugin.HtmlInlineElementConverters.Count; i++) {
                    var converter = plugin.HtmlInlineElementConverters[i];
                    if (converter == null) {
                        continue;
                    }

                    bool converterExists = false;
                    for (int j = 0; j < htmlInlineElementConverters.Count; j++) {
                        if (string.Equals(htmlInlineElementConverters[j].Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                            converterExists = true;
                            break;
                        }
                    }

                    if (!converterExists) {
                        htmlInlineElementConverters.Add(converter);
                    }
                }

                for (int i = 0; i < plugin.VisualElementRoundTripHints.Count; i++) {
                    var hint = plugin.VisualElementRoundTripHints[i];
                    if (hint == null) {
                        continue;
                    }

                    bool hintExists = false;
                    for (int j = 0; j < roundTripHints.Count; j++) {
                        if (string.Equals(roundTripHints[j].Id, hint.Id, StringComparison.OrdinalIgnoreCase)) {
                            hintExists = true;
                            break;
                        }
                    }

                    if (!hintExists) {
                        roundTripHints.Add(hint);
                    }
                }
            }
        }

        _plugins = normalizedPlugins.AsReadOnly();
        _readerDocumentTransforms = readerDocumentTransforms.AsReadOnly();
        _htmlDocumentTransforms = htmlDocumentTransforms.AsReadOnly();
        _rendererDocumentTransforms = rendererDocumentTransforms.AsReadOnly();
        _htmlElementBlockConverters = htmlElementBlockConverters.AsReadOnly();
        _htmlInlineElementConverters = htmlInlineElementConverters.AsReadOnly();
        _visualElementRoundTripHints = roundTripHints.AsReadOnly();
        _apply = apply ?? (_ => { });
        _applyReader = applyReader ?? (_ => { });
        Id = id.Trim();
        Name = name.Trim();
    }

    /// <summary>
    /// Stable feature-pack identifier used for idempotence and diagnostics.
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// Friendly feature-pack name used for diagnostics and documentation.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Reusable renderer plugins contributed by this feature pack.
    /// </summary>
    public IReadOnlyList<MarkdownRendererPlugin> Plugins => _plugins;

    /// <summary>
    /// Reader-side document transforms contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<IMarkdownDocumentTransform> ReaderDocumentTransforms => _readerDocumentTransforms;

    /// <summary>
    /// HTML-to-markdown document transforms contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<IMarkdownDocumentTransform> HtmlDocumentTransforms => _htmlDocumentTransforms;

    /// <summary>
    /// Renderer-side document transforms contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<IMarkdownDocumentTransform> RendererDocumentTransforms => _rendererDocumentTransforms;

    /// <summary>
    /// HTML-to-markdown custom element block converters contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<HtmlElementBlockConverter> HtmlElementBlockConverters => _htmlElementBlockConverters;

    /// <summary>
    /// HTML-to-markdown custom inline element converters contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<HtmlInlineElementConverter> HtmlInlineElementConverters => _htmlInlineElementConverters;

    /// <summary>
    /// Shared visual-host HTML round-trip hints contributed by the composed plugins in this feature pack.
    /// </summary>
    public IReadOnlyList<MarkdownVisualElementRoundTripHint> VisualElementRoundTripHints => _visualElementRoundTripHints;

    /// <summary>
    /// Applies the feature pack once to the supplied renderer options.
    /// </summary>
    public void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (!options.TryMarkFeaturePackApplied(Id)) {
            return;
        }

        for (int i = 0; i < _plugins.Count; i++) {
            options.ApplyPlugin(_plugins[i]);
        }

        options.ReaderOptions.ApplyFeaturePack(this);
        _apply(options);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the feature pack has already been applied.
    /// </summary>
    public bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasFeaturePackId(Id);
    }

    internal void ApplyReader(MarkdownReaderOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        for (int i = 0; i < _plugins.Count; i++) {
            options.ApplyPlugin(_plugins[i]);
        }

        _applyReader(options);
    }

    internal bool ReaderContractMatches(MarkdownReaderOptions options) {
        if (options == null) {
            return false;
        }

        if (_readerDocumentTransforms.Count == 0) {
            return false;
        }

        for (int i = 0; i < _readerDocumentTransforms.Count; i++) {
            if (!options.DocumentTransforms.Contains(_readerDocumentTransforms[i])) {
                return false;
            }
        }

        return true;
    }
}
