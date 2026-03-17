using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Bridges renderer plugins and feature packs into HTML-to-markdown ingestion configuration.
/// </summary>
public static class HtmlToMarkdownOptionsExtensions {
    /// <summary>
    /// Applies a renderer plugin's HTML-ingestion contract to HTML-to-markdown options.
    /// </summary>
    public static void ApplyPlugin(this HtmlToMarkdownOptions options, MarkdownRendererPlugin plugin) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (plugin == null) {
            throw new ArgumentNullException(nameof(plugin));
        }

        if (!options.TryMarkPluginApplied(plugin.Id)) {
            return;
        }

        AddDocumentTransforms(options.DocumentTransforms, plugin.HtmlDocumentTransforms);
        AddElementBlockConverters(options.ElementBlockConverters, plugin.HtmlElementBlockConverters);
        AddInlineElementConverters(options.InlineElementConverters, plugin.HtmlInlineElementConverters);
        AddRoundTripHints(options.VisualElementRoundTripHints, plugin.VisualElementRoundTripHints);
    }

    /// <summary>
    /// Returns <see langword="true"/> when a renderer plugin's HTML-ingestion contract
    /// has already been applied to HTML-to-markdown options.
    /// </summary>
    public static bool HasPlugin(this HtmlToMarkdownOptions options, MarkdownRendererPlugin plugin) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (plugin == null) {
            throw new ArgumentNullException(nameof(plugin));
        }

        return options.HasPluginId(plugin.Id)
            || (ContainsAllDocumentTransforms(options.DocumentTransforms, plugin.HtmlDocumentTransforms)
                && ContainsAllElementBlockConverters(options.ElementBlockConverters, plugin.HtmlElementBlockConverters)
                && ContainsAllInlineElementConverters(options.InlineElementConverters, plugin.HtmlInlineElementConverters)
                && ContainsAllHints(options.VisualElementRoundTripHints, plugin.VisualElementRoundTripHints));
    }

    /// <summary>
    /// Applies a renderer feature pack's HTML-ingestion contract to HTML-to-markdown options.
    /// </summary>
    public static void ApplyFeaturePack(this HtmlToMarkdownOptions options, MarkdownRendererFeaturePack featurePack) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (featurePack == null) {
            throw new ArgumentNullException(nameof(featurePack));
        }

        if (!options.TryMarkFeaturePackApplied(featurePack.Id)) {
            return;
        }

        for (int i = 0; i < featurePack.Plugins.Count; i++) {
            options.ApplyPlugin(featurePack.Plugins[i]);
        }

        AddDocumentTransforms(options.DocumentTransforms, featurePack.HtmlDocumentTransforms);
        AddElementBlockConverters(options.ElementBlockConverters, featurePack.HtmlElementBlockConverters);
        AddInlineElementConverters(options.InlineElementConverters, featurePack.HtmlInlineElementConverters);
        AddRoundTripHints(options.VisualElementRoundTripHints, featurePack.VisualElementRoundTripHints);
    }

    /// <summary>
    /// Returns <see langword="true"/> when a renderer feature pack's HTML-ingestion contract
    /// has already been applied to HTML-to-markdown options.
    /// </summary>
    public static bool HasFeaturePack(this HtmlToMarkdownOptions options, MarkdownRendererFeaturePack featurePack) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (featurePack == null) {
            throw new ArgumentNullException(nameof(featurePack));
        }

        return options.HasFeaturePackId(featurePack.Id)
            || (ContainsAllDocumentTransforms(options.DocumentTransforms, featurePack.HtmlDocumentTransforms)
                && ContainsAllElementBlockConverters(options.ElementBlockConverters, featurePack.HtmlElementBlockConverters)
                && ContainsAllInlineElementConverters(options.InlineElementConverters, featurePack.HtmlInlineElementConverters)
                && ContainsAllHints(options.VisualElementRoundTripHints, featurePack.VisualElementRoundTripHints));
    }

    private static void AddDocumentTransforms(ICollection<IMarkdownDocumentTransform> target, IReadOnlyList<IMarkdownDocumentTransform> transforms) {
        if (target == null || transforms == null || transforms.Count == 0) {
            return;
        }

        for (int i = 0; i < transforms.Count; i++) {
            var transform = transforms[i];
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

    private static void AddElementBlockConverters(ICollection<HtmlElementBlockConverter> target, IReadOnlyList<HtmlElementBlockConverter> converters) {
        if (target == null || converters == null || converters.Count == 0) {
            return;
        }

        for (int i = 0; i < converters.Count; i++) {
            var converter = converters[i];
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

    private static void AddInlineElementConverters(ICollection<HtmlInlineElementConverter> target, IReadOnlyList<HtmlInlineElementConverter> converters) {
        if (target == null || converters == null || converters.Count == 0) {
            return;
        }

        for (int i = 0; i < converters.Count; i++) {
            var converter = converters[i];
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

    private static void AddRoundTripHints(ICollection<MarkdownVisualElementRoundTripHint> target, IReadOnlyList<MarkdownVisualElementRoundTripHint> hints) {
        if (target == null || hints == null || hints.Count == 0) {
            return;
        }

        for (int i = 0; i < hints.Count; i++) {
            var hint = hints[i];
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

    private static bool ContainsAllHints(IReadOnlyList<MarkdownVisualElementRoundTripHint> target, IReadOnlyList<MarkdownVisualElementRoundTripHint> hints) {
        if (target == null || hints == null || hints.Count == 0) {
            return false;
        }

        for (int i = 0; i < hints.Count; i++) {
            var hint = hints[i];
            if (hint == null) {
                continue;
            }

            bool exists = false;
            for (int j = 0; j < target.Count; j++) {
                if (string.Equals(target[j].Id, hint.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                return false;
            }
        }

        return true;
    }

    private static bool ContainsAllDocumentTransforms(IReadOnlyList<IMarkdownDocumentTransform> target, IReadOnlyList<IMarkdownDocumentTransform> transforms) {
        if (target == null || transforms == null || transforms.Count == 0) {
            return false;
        }

        for (int i = 0; i < transforms.Count; i++) {
            var transform = transforms[i];
            if (transform == null) {
                continue;
            }

            bool exists = false;
            for (int j = 0; j < target.Count; j++) {
                if (ReferenceEquals(target[j], transform)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                return false;
            }
        }

        return true;
    }

    private static bool ContainsAllElementBlockConverters(IReadOnlyList<HtmlElementBlockConverter> target, IReadOnlyList<HtmlElementBlockConverter> converters) {
        if (target == null || converters == null || converters.Count == 0) {
            return false;
        }

        for (int i = 0; i < converters.Count; i++) {
            var converter = converters[i];
            if (converter == null) {
                continue;
            }

            bool exists = false;
            for (int j = 0; j < target.Count; j++) {
                if (string.Equals(target[j].Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                return false;
            }
        }

        return true;
    }

    private static bool ContainsAllInlineElementConverters(IReadOnlyList<HtmlInlineElementConverter> target, IReadOnlyList<HtmlInlineElementConverter> converters) {
        if (target == null || converters == null || converters.Count == 0) {
            return false;
        }

        for (int i = 0; i < converters.Count; i++) {
            var converter = converters[i];
            if (converter == null) {
                continue;
            }

            bool exists = false;
            for (int j = 0; j < target.Count; j++) {
                if (string.Equals(target[j].Id, converter.Id, StringComparison.OrdinalIgnoreCase)) {
                    exists = true;
                    break;
                }
            }

            if (!exists) {
                return false;
            }
        }

        return true;
    }
}
