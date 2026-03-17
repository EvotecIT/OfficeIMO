using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Reusable fenced-block renderer plugin that can register one or more renderer extensions on top of
/// <see cref="OfficeIMO.Markdown"/> and the host HTML shell.
/// </summary>
public sealed class MarkdownRendererPlugin {
    private readonly IReadOnlyList<RendererRegistration> _registrations;
    private readonly IReadOnlyList<MarkdownFenceOptionSchema> _fenceOptionSchemas;
    private readonly IReadOnlyList<IMarkdownDocumentTransform> _readerDocumentTransforms;
    private readonly IReadOnlyList<IMarkdownDocumentTransform> _htmlDocumentTransforms;
    private readonly IReadOnlyList<HtmlElementBlockConverter> _htmlElementBlockConverters;
    private readonly IReadOnlyList<HtmlInlineElementConverter> _htmlInlineElementConverters;
    private readonly IReadOnlyList<MarkdownVisualElementRoundTripHint> _visualElementRoundTripHints;
    private readonly Action<MarkdownRendererOptions> _apply;
    private readonly Action<MarkdownReaderOptions> _applyReader;

    /// <summary>
    /// Creates a plugin from one or more fenced-block renderer factories.
    /// </summary>
    public MarkdownRendererPlugin(
        string name,
        IEnumerable<Func<MarkdownFencedCodeBlockRenderer>> rendererFactories,
        IEnumerable<MarkdownFenceOptionSchema>? fenceOptionSchemas = null,
        IEnumerable<IMarkdownDocumentTransform>? readerDocumentTransforms = null,
        IEnumerable<IMarkdownDocumentTransform>? htmlDocumentTransforms = null,
        IEnumerable<HtmlElementBlockConverter>? htmlElementBlockConverters = null,
        IEnumerable<HtmlInlineElementConverter>? htmlInlineElementConverters = null,
        IEnumerable<MarkdownVisualElementRoundTripHint>? visualElementRoundTripHints = null,
        Action<MarkdownRendererOptions>? apply = null,
        Action<MarkdownReaderOptions>? applyReader = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Plugin name is required.", nameof(name));
        }

        if (rendererFactories == null) {
            throw new ArgumentNullException(nameof(rendererFactories));
        }

        var registrations = new List<RendererRegistration>();
        foreach (var factory in rendererFactories) {
            if (factory == null) {
                continue;
            }

            var renderer = factory();
            registrations.Add(new RendererRegistration(renderer.Languages, factory));
        }

        if (registrations.Count == 0) {
            throw new ArgumentException("At least one fenced code block renderer factory is required.", nameof(rendererFactories));
        }

        Name = name.Trim();
        Id = Name;
        _registrations = registrations;
        _fenceOptionSchemas = NormalizeSchemas(fenceOptionSchemas);
        _readerDocumentTransforms = NormalizeDocumentTransforms(readerDocumentTransforms);
        _htmlDocumentTransforms = NormalizeHtmlDocumentTransforms(htmlDocumentTransforms);
        _htmlElementBlockConverters = NormalizeHtmlElementBlockConverters(htmlElementBlockConverters);
        _htmlInlineElementConverters = NormalizeHtmlInlineElementConverters(htmlInlineElementConverters);
        _visualElementRoundTripHints = NormalizeRoundTripHints(visualElementRoundTripHints);
        _apply = apply ?? (_ => { });
        _applyReader = applyReader ?? (_ => { });
    }

    /// <summary>
    /// Creates a plugin by composing one or more existing plugins and optional extra schemas.
    /// </summary>
    public MarkdownRendererPlugin(
        string name,
        IEnumerable<MarkdownRendererPlugin> plugins,
        IEnumerable<MarkdownFenceOptionSchema>? fenceOptionSchemas = null,
        IEnumerable<IMarkdownDocumentTransform>? readerDocumentTransforms = null,
        IEnumerable<IMarkdownDocumentTransform>? htmlDocumentTransforms = null,
        IEnumerable<HtmlElementBlockConverter>? htmlElementBlockConverters = null,
        IEnumerable<HtmlInlineElementConverter>? htmlInlineElementConverters = null,
        IEnumerable<MarkdownVisualElementRoundTripHint>? visualElementRoundTripHints = null,
        Action<MarkdownRendererOptions>? apply = null,
        Action<MarkdownReaderOptions>? applyReader = null) {
        if (string.IsNullOrWhiteSpace(name)) {
            throw new ArgumentException("Plugin name is required.", nameof(name));
        }

        if (plugins == null) {
            throw new ArgumentNullException(nameof(plugins));
        }

        var normalizedPlugins = new List<MarkdownRendererPlugin>();
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
            }
        }

        var registrations = new List<RendererRegistration>();
        var schemas = new List<MarkdownFenceOptionSchema>();
        var readerDocumentTransformsList = new List<IMarkdownDocumentTransform>();
        var htmlDocumentTransformsList = new List<IMarkdownDocumentTransform>();
        var htmlElementBlockConvertersList = new List<HtmlElementBlockConverter>();
        var htmlInlineElementConvertersList = new List<HtmlInlineElementConverter>();
        var roundTripHints = new List<MarkdownVisualElementRoundTripHint>();
        var childApplyActions = new List<Action<MarkdownRendererOptions>>();
        var childReaderApplyActions = new List<Action<MarkdownReaderOptions>>();
        foreach (var plugin in normalizedPlugins) {
            AddRegistrations(registrations, plugin._registrations);
            AddSchemas(schemas, plugin._fenceOptionSchemas);
            AddDocumentTransforms(readerDocumentTransformsList, plugin._readerDocumentTransforms);
            AddHtmlDocumentTransforms(htmlDocumentTransformsList, plugin._htmlDocumentTransforms);
            AddHtmlElementBlockConverters(htmlElementBlockConvertersList, plugin._htmlElementBlockConverters);
            AddHtmlInlineElementConverters(htmlInlineElementConvertersList, plugin._htmlInlineElementConverters);
            AddRoundTripHints(roundTripHints, plugin._visualElementRoundTripHints);
            childApplyActions.Add(plugin._apply);
            childReaderApplyActions.Add(plugin._applyReader);
        }

        AddSchemas(schemas, NormalizeSchemas(fenceOptionSchemas));
        AddDocumentTransforms(readerDocumentTransformsList, NormalizeDocumentTransforms(readerDocumentTransforms));
        AddHtmlDocumentTransforms(htmlDocumentTransformsList, NormalizeHtmlDocumentTransforms(htmlDocumentTransforms));
        AddHtmlElementBlockConverters(htmlElementBlockConvertersList, NormalizeHtmlElementBlockConverters(htmlElementBlockConverters));
        AddHtmlInlineElementConverters(htmlInlineElementConvertersList, NormalizeHtmlInlineElementConverters(htmlInlineElementConverters));
        AddRoundTripHints(roundTripHints, NormalizeRoundTripHints(visualElementRoundTripHints));

        if (registrations.Count == 0) {
            throw new ArgumentException("At least one composed plugin with fenced code block registrations is required.", nameof(plugins));
        }

        Name = name.Trim();
        Id = Name;
        _registrations = registrations.AsReadOnly();
        _fenceOptionSchemas = schemas.AsReadOnly();
        _readerDocumentTransforms = readerDocumentTransformsList.AsReadOnly();
        _htmlDocumentTransforms = htmlDocumentTransformsList.AsReadOnly();
        _htmlElementBlockConverters = htmlElementBlockConvertersList.AsReadOnly();
        _htmlInlineElementConverters = htmlInlineElementConvertersList.AsReadOnly();
        _visualElementRoundTripHints = roundTripHints.AsReadOnly();
        _apply = options => {
            for (int i = 0; i < childApplyActions.Count; i++) {
                childApplyActions[i](options);
            }

            apply?.Invoke(options);
        };
        _applyReader = options => {
            for (int i = 0; i < childReaderApplyActions.Count; i++) {
                childReaderApplyActions[i](options);
            }

            applyReader?.Invoke(options);
        };
    }

    /// <summary>
    /// Friendly plugin name used for diagnostics and documentation.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Stable plugin identifier used for idempotence and diagnostics.
    /// </summary>
    public string Id { get; }

    /// <summary>
    /// All fenced block languages covered by this plugin.
    /// </summary>
    public IReadOnlyList<string> Languages {
        get {
            var values = new List<string>();
            for (int i = 0; i < _registrations.Count; i++) {
                AddLanguages(values, _registrations[i].Languages);
            }

            return values;
        }
    }

    /// <summary>
    /// Fence option schemas contributed by this plugin.
    /// </summary>
    public IReadOnlyList<MarkdownFenceOptionSchema> FenceOptionSchemas => _fenceOptionSchemas;

    /// <summary>
    /// Reader-side document transforms contributed by this plugin.
    /// Hosts can register these on <see cref="MarkdownReaderOptions.DocumentTransforms"/>
    /// when they want source parsing aligned with the plugin contract.
    /// </summary>
    public IReadOnlyList<IMarkdownDocumentTransform> ReaderDocumentTransforms => _readerDocumentTransforms;

    /// <summary>
    /// HTML-to-markdown document transforms contributed by this plugin.
    /// Hosts can register these on <c>HtmlToMarkdownOptions.DocumentTransforms</c>
    /// when they want HTML ingestion aligned with the plugin contract.
    /// </summary>
    public IReadOnlyList<IMarkdownDocumentTransform> HtmlDocumentTransforms => _htmlDocumentTransforms;

    /// <summary>
    /// HTML-to-markdown custom element block converters contributed by this plugin.
    /// Hosts can register these on <c>HtmlToMarkdownOptions.ElementBlockConverters</c>
    /// when they want HTML element decoding aligned with the plugin contract.
    /// </summary>
    public IReadOnlyList<HtmlElementBlockConverter> HtmlElementBlockConverters => _htmlElementBlockConverters;

    /// <summary>
    /// HTML-to-markdown custom inline element converters contributed by this plugin.
    /// Hosts can register these on <c>HtmlToMarkdownOptions.InlineElementConverters</c>
    /// when they want inline HTML decoding aligned with the plugin contract.
    /// </summary>
    public IReadOnlyList<HtmlInlineElementConverter> HtmlInlineElementConverters => _htmlInlineElementConverters;

    /// <summary>
    /// Shared visual-host HTML round-trip hints contributed by this plugin.
    /// Hosts can register these on <c>HtmlToMarkdownOptions.VisualElementRoundTripHints</c>
    /// when they want HTML-to-AST recovery aligned with the plugin contract.
    /// </summary>
    public IReadOnlyList<MarkdownVisualElementRoundTripHint> VisualElementRoundTripHints => _visualElementRoundTripHints;

    /// <summary>
    /// Applies the plugin to the supplied renderer options without duplicating existing language registrations.
    /// </summary>
    public void Apply(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (!options.TryMarkPluginApplied(Id)) {
            return;
        }

        for (int i = 0; i < _fenceOptionSchemas.Count; i++) {
            options.ApplyFenceOptionSchema(_fenceOptionSchemas[i]);
        }

        options.ReaderOptions.ApplyPlugin(this);

        for (int i = 0; i < _registrations.Count; i++) {
            var registration = _registrations[i];
            if (!HasAnyLanguage(options.FencedCodeBlockRenderers, registration.Languages)) {
                options.FencedCodeBlockRenderers.Add(registration.Factory());
            }
        }

        _apply(options);
    }

    /// <summary>
    /// Returns <see langword="true"/> when any of this plugin's fenced block languages are already registered.
    /// </summary>
    public bool IsApplied(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (options.HasPluginId(Id)) {
            return true;
        }

        for (int i = 0; i < _registrations.Count; i++) {
            if (HasAnyLanguage(options.FencedCodeBlockRenderers, _registrations[i].Languages)) {
                return true;
            }
        }

        return false;
    }

    private static void AddLanguages(List<string> target, IReadOnlyList<string> languages) {
        if (languages == null || languages.Count == 0) {
            return;
        }

        for (int i = 0; i < languages.Count; i++) {
            var candidate = languages[i];
            if (string.IsNullOrWhiteSpace(candidate)) {
                continue;
            }

            bool exists = false;
            for (int j = 0; j < target.Count; j++) {
                if (string.Equals(target[j], candidate, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(candidate);
            }
        }
    }

    private static bool HasAnyLanguage(IReadOnlyList<MarkdownFencedCodeBlockRenderer> renderers, IReadOnlyList<string> languages) {
        if (renderers == null || languages == null || languages.Count == 0) {
            return false;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            for (int j = 0; j < languages.Count; j++) {
                var candidate = languages[j];
                if (string.IsNullOrWhiteSpace(candidate)) {
                    continue;
                }

                if (RendererHandlesLanguage(renderer, candidate)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static bool RendererHandlesLanguage(MarkdownFencedCodeBlockRenderer renderer, string language) {
        var languages = renderer.Languages;
        if (languages == null || languages.Count == 0) {
            return false;
        }

        for (int i = 0; i < languages.Count; i++) {
            if (string.Equals(languages[i], language, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static IReadOnlyList<MarkdownFenceOptionSchema> NormalizeSchemas(IEnumerable<MarkdownFenceOptionSchema>? schemas) {
        var normalized = new List<MarkdownFenceOptionSchema>();
        AddSchemas(normalized, schemas);
        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<IMarkdownDocumentTransform> NormalizeDocumentTransforms(IEnumerable<IMarkdownDocumentTransform>? transforms) {
        var normalized = new List<IMarkdownDocumentTransform>();
        AddDocumentTransforms(normalized, transforms);
        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<IMarkdownDocumentTransform> NormalizeHtmlDocumentTransforms(IEnumerable<IMarkdownDocumentTransform>? transforms) {
        var normalized = new List<IMarkdownDocumentTransform>();
        AddHtmlDocumentTransforms(normalized, transforms);
        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<HtmlElementBlockConverter> NormalizeHtmlElementBlockConverters(IEnumerable<HtmlElementBlockConverter>? converters) {
        var normalized = new List<HtmlElementBlockConverter>();
        AddHtmlElementBlockConverters(normalized, converters);
        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<HtmlInlineElementConverter> NormalizeHtmlInlineElementConverters(IEnumerable<HtmlInlineElementConverter>? converters) {
        var normalized = new List<HtmlInlineElementConverter>();
        AddHtmlInlineElementConverters(normalized, converters);
        return normalized.AsReadOnly();
    }

    private static IReadOnlyList<MarkdownVisualElementRoundTripHint> NormalizeRoundTripHints(IEnumerable<MarkdownVisualElementRoundTripHint>? hints) {
        var normalized = new List<MarkdownVisualElementRoundTripHint>();
        AddRoundTripHints(normalized, hints);
        return normalized.AsReadOnly();
    }

    private static void AddSchemas(ICollection<MarkdownFenceOptionSchema> target, IEnumerable<MarkdownFenceOptionSchema>? schemas) {
        if (schemas == null) {
            return;
        }

        foreach (var schema in schemas) {
            if (schema == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (string.Equals(existing.Id, schema.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(schema);
            }
        }
    }

    private static void AddRoundTripHints(ICollection<MarkdownVisualElementRoundTripHint> target, IEnumerable<MarkdownVisualElementRoundTripHint>? hints) {
        if (hints == null) {
            return;
        }

        foreach (var hint in hints) {
            if (hint == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (string.Equals(existing.Id, hint.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(hint);
            }
        }
    }

    private static void AddDocumentTransforms(ICollection<IMarkdownDocumentTransform> target, IEnumerable<IMarkdownDocumentTransform>? transforms) {
        if (transforms == null) {
            return;
        }

        foreach (var transform in transforms) {
            if (transform == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (ReferenceEquals(existing, transform)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(transform);
            }
        }
    }

    private static void AddHtmlDocumentTransforms(ICollection<IMarkdownDocumentTransform> target, IEnumerable<IMarkdownDocumentTransform>? transforms) {
        if (transforms == null) {
            return;
        }

        foreach (var transform in transforms) {
            if (transform == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (ReferenceEquals(existing, transform)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(transform);
            }
        }
    }

    private static void AddHtmlElementBlockConverters(ICollection<HtmlElementBlockConverter> target, IEnumerable<HtmlElementBlockConverter>? converters) {
        if (converters == null) {
            return;
        }

        foreach (var converter in converters) {
            if (converter == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (string.Equals(existing.Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(converter);
            }
        }
    }

    private static void AddHtmlInlineElementConverters(ICollection<HtmlInlineElementConverter> target, IEnumerable<HtmlInlineElementConverter>? converters) {
        if (converters == null) {
            return;
        }

        foreach (var converter in converters) {
            if (converter == null) {
                continue;
            }

            bool exists = false;
            foreach (var existing in target) {
                if (string.Equals(existing.Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                target.Add(converter);
            }
        }
    }

    private static void AddRegistrations(ICollection<RendererRegistration> target, IEnumerable<RendererRegistration> registrations) {
        foreach (var registration in registrations) {
            if (registration == null) {
                continue;
            }

            target.Add(registration);
        }
    }

    private sealed class RendererRegistration {
        public RendererRegistration(IReadOnlyList<string> languages, Func<MarkdownFencedCodeBlockRenderer> factory) {
            Languages = languages ?? Array.Empty<string>();
            Factory = factory ?? throw new ArgumentNullException(nameof(factory));
        }

        public IReadOnlyList<string> Languages { get; }
        public Func<MarkdownFencedCodeBlockRenderer> Factory { get; }
    }

    internal void ApplyReader(MarkdownReaderOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        for (int i = 0; i < _readerDocumentTransforms.Count; i++) {
            var transform = _readerDocumentTransforms[i];
            if (transform != null && !options.DocumentTransforms.Contains(transform)) {
                options.DocumentTransforms.Add(transform);
            }
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
