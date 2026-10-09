using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

/// <summary>
/// Produces stable, policy-aware normalized HTML for OfficeIMO conversion workflows.
/// </summary>
public static partial class HtmlNormalizer {
    private static readonly HashSet<string> BooleanAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "allowfullscreen", "async", "autofocus", "autoplay", "checked", "controls", "default", "defer", "disabled",
        "formnovalidate", "hidden", "loop", "multiple", "muted", "nomodule", "novalidate", "open", "readonly",
        "required", "reversed", "selected"
    };
    private static readonly HashSet<string> VoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "area", "base", "br", "col", "embed", "hr", "img", "input", "link", "meta", "source", "track", "wbr"
    };
    private static readonly HashSet<string> SkippedElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "script", "template"
    };
    private static readonly HashSet<string> UrlAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "action", "background", "cite", "data-original", "data-original-src", "data-lazy-src", "data-poster", "data-src", "formaction", "href", "poster", "src", "xlink:href"
    };
    private static readonly HashSet<string> LazyUrlAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "data-original", "data-original-src", "data-lazy-src", "data-poster", "data-src"
    };
    private static readonly HashSet<string> SrcSetAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "data-original-srcset", "data-lazy-srcset", "data-srcset", "imagesrcset", "srcset"
    };
    private static readonly HashSet<string> LazySrcSetAttributes = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "data-original-srcset", "data-lazy-srcset", "data-srcset"
    };
    private static readonly char[] WhitespaceSeparators = { ' ', '\t', '\r', '\n', '\f' };

    /// <summary>
    /// Parses and normalizes raw HTML.
    /// </summary>
    public static string Normalize(string html, HtmlNormalizationOptions? options = null) {
        if (html == null) {
            throw new ArgumentNullException(nameof(html));
        }

        HtmlNormalizationOptions resolved = CopyOptions(options ?? new HtmlNormalizationOptions());
        resolved.Limits.Validate();
        IHtmlDocument document = HtmlConversionDocument.ParseSourceDocumentForAnalysis(html, resolved.Limits);
        return NormalizeDocument(document, resolved, 0);
    }

    /// <summary>
    /// Normalizes an already parsed HTML document.
    /// </summary>
    public static string Normalize(Dom.HtmlDocument document, HtmlNormalizationOptions? options = null) {
        HtmlNormalizationOptions resolved = CopyOptions(options ?? new HtmlNormalizationOptions());
        Dom.HtmlDocument snapshot = HtmlConversionInputGuard.CaptureOwnedTree(document, resolved.Limits, CancellationToken.None);
        return Normalize(NativeDomBridge.GetNativeDocument(snapshot), resolved);
    }

    internal static string Normalize(IHtmlDocument document, HtmlNormalizationOptions? options = null, CancellationToken cancellationToken = default) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        HtmlNormalizationOptions resolved = CopyOptions(options ?? new HtmlNormalizationOptions());
        resolved.Limits.Validate();
        HtmlConversionInputGuard.ValidateDocument(document, resolved.Limits, cancellationToken);
        return NormalizeDocument(document, resolved, 0, cancellationToken);
    }

    /// <summary>Normalizes one already parsed element for a raw-fragment target without reparsing it.</summary>
    internal static string NormalizePreparedElement(IElement element, HtmlNormalizationOptions options, CancellationToken cancellationToken = default) {
        if (element == null) throw new ArgumentNullException(nameof(element));
        HtmlNormalizationOptions resolved = CopyOptions(options ?? throw new ArgumentNullException(nameof(options)));
        resolved.Limits.Validate();
        var builder = new StringBuilder();
        AppendElement(builder, element, resolved, 0, cancellationToken: cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString();
    }

    /// <summary>
    /// Removes executable element payloads and inline event handlers from a prepared adapter DOM
    /// without resolving URL attributes that the target adapter still needs to diagnose.
    /// </summary>
    internal static void SanitizePreparedDocumentStructure(IHtmlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        foreach (IElement element in document.QuerySelectorAll("*")) {
            foreach (IAttr attribute in element.Attributes
                .Where(attribute => attribute.Name.StartsWith("on", StringComparison.OrdinalIgnoreCase))
                .ToArray()) {
                element.RemoveAttribute(attribute.Name);
            }

            if (SkippedElements.Contains(element.LocalName)) element.TextContent = string.Empty;
        }
    }

    private static string NormalizeDocument(IHtmlDocument document, HtmlNormalizationOptions options, int srcDocDepth, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        INode root = HtmlDocumentParser.GetConversionRoot(document, options.UseBodyContentsOnly);

        var builder = new StringBuilder();
        if (!options.UseBodyContentsOnly && root is IElement documentElement) {
            AppendElement(builder, documentElement, options, srcDocDepth, cancellationToken: cancellationToken);
        } else {
            foreach (INode child in root.ChildNodes) {
                AppendNode(builder, child, options, srcDocDepth, preserveWhitespace: false, cancellationToken);
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        return builder.ToString().Trim();
    }

    private static void AppendNode(StringBuilder builder, INode node, HtmlNormalizationOptions options, int srcDocDepth, bool preserveWhitespace = false, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (node.NodeType == NodeType.Text) {
            AppendText(builder, node.TextContent, options, preserveWhitespace, cancellationToken);
            return;
        }

        if (node.NodeType == NodeType.Comment) {
            if (options.PreserveComments) {
                builder.Append("<!--").Append(node.TextContent).Append("-->");
            }

            return;
        }

        if (node is IElement element) {
            AppendElement(builder, element, options, srcDocDepth, preserveWhitespace, cancellationToken);
        }
    }

    private static void AppendElement(StringBuilder builder, IElement element, HtmlNormalizationOptions options, int srcDocDepth, bool preserveWhitespace = false, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        string name = element.TagName.ToLowerInvariant();
        if (SkippedElements.Contains(name)) {
            if (options.PreserveSkippedElementMarkers) {
                builder.Append('<').Append(name).Append("></").Append(name).Append('>');
            }

            return;
        }

        if (name == "style" && !options.PreserveStyleElements) {
            return;
        }

        string emittedName = IsForeignContent(element) ? element.TagName : name;
        builder.Append('<').Append(emittedName);
        foreach (KeyValuePair<string, string> attribute in NormalizeAttributes(element, options, srcDocDepth, cancellationToken)) {
            builder.Append(' ').Append(attribute.Key);
            if (!BooleanAttributes.Contains(attribute.Key) || !IsBooleanValue(attribute.Key, attribute.Value)) {
                builder.Append("=\"").Append(WebUtility.HtmlEncode(attribute.Value)).Append('"');
            }
        }

        builder.Append('>');
        if (!VoidElements.Contains(name)) {
            if (name == "style" && options.PreserveStyleElements) {
                string styleText = NormalizeCssUrls(element.TextContent, options.BaseUri, GetResourceUrlPolicy(options), cancellationToken);
                if (styleText.Length > 0) {
                    builder.Append(EscapeRawTextElementContent(styleText, "style"));
                }
            } else {
                bool childPreserveWhitespace = preserveWhitespace || IsPreformattedElement(name);
                foreach (INode child in element.ChildNodes) {
                    AppendNode(builder, child, options, srcDocDepth, childPreserveWhitespace, cancellationToken);
                }
            }

            builder.Append("</").Append(emittedName).Append('>');
        }
    }

    private static IReadOnlyList<KeyValuePair<string, string>> NormalizeAttributes(IElement element, HtmlNormalizationOptions options, int srcDocDepth, CancellationToken cancellationToken) {
        var attributes = new List<KeyValuePair<string, string>>();
        bool preserveAttributeCasing = IsForeignContent(element);
        foreach (IAttr attribute in element.Attributes) {
            cancellationToken.ThrowIfCancellationRequested();
            string qualifiedName = HtmlDocumentParser.GetQualifiedAttributeName(attribute);
            string name = qualifiedName.ToLowerInvariant();
            if (options.RemoveEventHandlerAttributes && name.StartsWith("on", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            string value = NormalizeAttributeValue(element, name, attribute.Value, options, srcDocDepth, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            if (IsUrlAttribute(element, name) || IsSrcSetAttribute(element, name)) {
                if (string.IsNullOrWhiteSpace(value) && !ShouldPreserveEmptyUrlAttribute(element, name, attribute.Value)) {
                    continue;
                }
            }

            string emittedName = preserveAttributeCasing ? qualifiedName : name;
            attributes.Add(new KeyValuePair<string, string>(emittedName, value));
        }

        return attributes
            .OrderBy(pair => AttributeOrder(pair.Key))
            .ThenBy(pair => pair.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string NormalizeAttributeValue(IElement element, string name, string value, HtmlNormalizationOptions options, int srcDocDepth, CancellationToken cancellationToken) {
        if (BooleanAttributes.Contains(name) && IsBooleanValue(name, value)) {
            return name;
        }

        if (string.Equals(name, "srcdoc", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeSrcDoc(value, options, srcDocDepth, cancellationToken);
        }

        if (IsSrcSetAttribute(element, name)) {
            return HtmlImageSourceResolver.ResolveNormalizedSrcSet(
                value,
                options.BaseUri,
                GetResourceUrlPolicy(options),
                options.MaxResponsiveImageCandidates,
                cancellationToken);
        }

        if (IsMetaRefreshContentAttribute(element, name)) {
            return NormalizeMetaRefreshContent(value, options, cancellationToken);
        }

        if (IsUrlAttribute(element, name)) {
            HtmlUrlPolicy attributePolicy = GetAttributeUrlPolicy(element, name, options);
            if (string.IsNullOrWhiteSpace(value) && ShouldPreserveEmptyUrlAttribute(element, name, value)) {
                return HtmlUrlPolicyEvaluator.ResolveUrl(options.BaseUri?.AbsoluteUri, null, attributePolicy);
            }

            Uri? baseUri = string.Equals(element.TagName, "base", StringComparison.OrdinalIgnoreCase)
                && string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)
                && options.BaseElementBaseUri != null
                    ? options.BaseElementBaseUri
                    : options.BaseUri;
            return HtmlUrlPolicyEvaluator.ResolveUrl(value, baseUri, attributePolicy);
        }

        if (string.Equals(name, "style", StringComparison.OrdinalIgnoreCase)) {
            return NormalizeCssUrls(value, options.BaseUri, GetResourceUrlPolicy(options), cancellationToken).Trim();
        }

        if (string.Equals(name, "class", StringComparison.OrdinalIgnoreCase)) {
            return string.Join(" ", value.Split(WhitespaceSeparators, StringSplitOptions.RemoveEmptyEntries));
        }

        return value;
    }

    private static string NormalizeSrcDoc(string value, HtmlNormalizationOptions options, int srcDocDepth, CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(value) || srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) {
            return string.Empty;
        }

        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(value, cancellationToken);
        HtmlNormalizationOptions nestedOptions = CopyOptions(options);
        nestedOptions.BaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, options.BaseUri);
        nestedOptions.UseBodyContentsOnly = true;
        return NormalizeDocument(nested, nestedOptions, srcDocDepth + 1, cancellationToken);
    }

    private static HtmlNormalizationOptions CopyOptions(HtmlNormalizationOptions options) {
        HtmlConversionLimits limits = (options.Limits ?? HtmlConversionLimits.CreateUntrustedProfile()).Clone();
        return new HtmlNormalizationOptions {
            BaseUri = options.BaseUri,
            BaseElementBaseUri = options.BaseElementBaseUri,
            UrlPolicy = (options.UrlPolicy ?? HtmlUrlPolicy.CreateOfficeIMOProfile()).Clone(),
            ResourceUrlPolicy = (options.ResourceUrlPolicy ?? HtmlResourceUrlPolicy.Create(options.UrlPolicy)).Clone(),
            Limits = limits,
            MaxResponsiveImageCandidates = HtmlConversionLimits.Minimum(options.MaxResponsiveImageCandidates, limits.MaxResponsiveImageCandidates),
            UseBodyContentsOnly = options.UseBodyContentsOnly,
            PreserveComments = options.PreserveComments,
            PreserveSkippedElementMarkers = options.PreserveSkippedElementMarkers,
            PreserveStyleElements = options.PreserveStyleElements,
            RemoveEventHandlerAttributes = options.RemoveEventHandlerAttributes,
            CollapseTextWhitespace = options.CollapseTextWhitespace
        };
    }

    private static void AppendText(StringBuilder builder, string? text, HtmlNormalizationOptions options, bool preserveWhitespace, CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        string value = options.CollapseTextWhitespace && !preserveWhitespace
            ? CollapseWhitespaceRuns(text!, cancellationToken)
            : text!;
        if (value.Length > 0) {
            if (options.CollapseTextWhitespace && !preserveWhitespace) {
                if (value == " ") {
                    if (builder.Length == 0 || char.IsWhiteSpace(builder[builder.Length - 1])) {
                        return;
                    }
                } else if (value[0] == ' ' && (builder.Length == 0 || char.IsWhiteSpace(builder[builder.Length - 1]))) {
                    value = value.TrimStart();
                }
            }

            builder.Append(WebUtility.HtmlEncode(value));
        }
    }

    private static bool IsPreformattedElement(string name) {
        return string.Equals(name, "pre", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "textarea", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsUrlAttribute(IElement element, string name) {
        if (LazyUrlAttributes.Contains(name)) {
            return IsLazyUrlElement(element, name);
        }

        if (UrlAttributes.Contains(name)) {
            return true;
        }

        return string.Equals(name, "data", StringComparison.OrdinalIgnoreCase)
            && string.Equals(element.TagName, "object", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSrcSetAttribute(IElement element, string name) {
        if (!SrcSetAttributes.Contains(name)) {
            return false;
        }

        if (LazySrcSetAttributes.Contains(name)) {
            return IsImageSourceElement(element);
        }

        if (string.Equals(name, "imagesrcset", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(element.TagName, "link", StringComparison.OrdinalIgnoreCase);
        }

        return IsImageSourceElement(element);
    }

    private static bool IsLazyUrlElement(IElement element, string name) {
        string tagName = element.TagName.ToLowerInvariant();
        if (string.Equals(name, "data-poster", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(tagName, "video", StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(name, "data-src", StringComparison.OrdinalIgnoreCase)
            && (string.Equals(tagName, "video", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "audio", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "track", StringComparison.OrdinalIgnoreCase))) {
            return true;
        }

        return IsImageSourceElement(element)
            || (string.Equals(tagName, "input", StringComparison.OrdinalIgnoreCase)
                && string.Equals(
                    HtmlFormControlSemantics.GetEffectiveType("input", element.GetAttribute("type")),
                    "image",
                    StringComparison.Ordinal));
    }

    private static bool IsImageSourceElement(IElement element) {
        string tagName = element.TagName.ToLowerInvariant();
        return string.Equals(tagName, "img", StringComparison.OrdinalIgnoreCase)
            || string.Equals(tagName, "image", StringComparison.OrdinalIgnoreCase)
            || string.Equals(tagName, "source", StringComparison.OrdinalIgnoreCase);
    }

    private static bool ShouldPreserveEmptyUrlAttribute(IElement element, string name, string value) {
        if (!string.IsNullOrWhiteSpace(value)) {
            return false;
        }

        string tagName = element.TagName.ToLowerInvariant();
        return (string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)
                && (string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(tagName, "area", StringComparison.OrdinalIgnoreCase)))
            || (string.Equals(name, "action", StringComparison.OrdinalIgnoreCase)
                && string.Equals(tagName, "form", StringComparison.OrdinalIgnoreCase))
            || (string.Equals(name, "formaction", StringComparison.OrdinalIgnoreCase)
                && (string.Equals(tagName, "button", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(tagName, "input", StringComparison.OrdinalIgnoreCase)));
    }

    private static bool IsMetaRefreshContentAttribute(IElement element, string name) {
        return string.Equals(name, "content", StringComparison.OrdinalIgnoreCase)
            && string.Equals(element.TagName, "meta", StringComparison.OrdinalIgnoreCase)
            && string.Equals(element.GetAttribute("http-equiv"), "refresh", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeMetaRefreshContent(string content, HtmlNormalizationOptions options, CancellationToken cancellationToken) {
        List<string> parameters = SplitMetaRefreshParameters(content).ToList();
        if (parameters.Count == 0) {
            return content;
        }

        var normalizedParameters = new List<string> {
            parameters[0]
        };
        bool changedUrl = false;
        for (int i = 1; i < parameters.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            string parameter = parameters[i];
            int separator = parameter.IndexOf('=');
            if (separator <= 0 || !string.Equals(parameter.Substring(0, separator).Trim(), "url", StringComparison.OrdinalIgnoreCase)) {
                normalizedParameters.Add(parameter);
                continue;
            }

            changedUrl = true;
            string source = parameter.Substring(separator + 1).Trim();
            if (source.Length > 1 && ((source[0] == '"' && source[source.Length - 1] == '"') || (source[0] == '\'' && source[source.Length - 1] == '\''))) {
                source = source.Substring(1, source.Length - 2).Trim();
            }

            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, options.BaseUri, options.UrlPolicy);
            cancellationToken.ThrowIfCancellationRequested();
            if (!string.IsNullOrWhiteSpace(resolved)) {
                normalizedParameters.Add("url=" + resolved);
            }
        }

        return changedUrl
            ? string.Join("; ", normalizedParameters.Where(parameter => parameter.Length > 0))
            : content;
    }

    private static IEnumerable<string> SplitMetaRefreshParameters(string content) {
        int start = 0;
        char quote = '\0';
        for (int i = 0; i < content.Length; i++) {
            char current = content[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(content, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == ';') {
                yield return content.Substring(start, i - start).Trim();
                start = i + 1;
            }
        }

        yield return content.Substring(start).Trim();
    }

    private static HtmlUrlPolicy GetAttributeUrlPolicy(IElement element, string name, HtmlNormalizationOptions options) {
        return IsHyperlinkUrlAttribute(element, name)
            ? options.UrlPolicy
            : GetResourceUrlPolicy(options);
    }

    private static HtmlUrlPolicy GetResourceUrlPolicy(HtmlNormalizationOptions options) {
        return options.ResourceUrlPolicy ?? HtmlResourceUrlPolicy.Create(options.UrlPolicy);
    }

    private static bool IsHyperlinkUrlAttribute(IElement element, string name) {
        string tagName = element.TagName.ToLowerInvariant();
        if (string.Equals(name, "href", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(tagName, "a", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "area", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "base", StringComparison.OrdinalIgnoreCase);
        }

        if (string.Equals(name, "cite", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        if (string.Equals(name, "action", StringComparison.OrdinalIgnoreCase)) {
            return string.Equals(tagName, "form", StringComparison.OrdinalIgnoreCase);
        }

        return string.Equals(name, "formaction", StringComparison.OrdinalIgnoreCase)
            && (string.Equals(tagName, "button", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "input", StringComparison.OrdinalIgnoreCase));
    }

    private static string CollapseWhitespaceRuns(string text, CancellationToken cancellationToken) {
        var builder = new StringBuilder(text.Length);
        bool inWhitespace = false;
        for (int i = 0; i < text.Length; i++) {
            if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            char current = text[i];
            if (char.IsWhiteSpace(current)) {
                if (!inWhitespace) {
                    builder.Append(' ');
                    inWhitespace = true;
                }

                continue;
            }

            builder.Append(current);
            inWhitespace = false;
        }

        return builder.ToString();
    }

    private static bool IsForeignContent(IElement element) {
        IElement? current = element;
        while (current != null) {
            if (string.Equals(current.TagName, "svg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(current.TagName, "math", StringComparison.OrdinalIgnoreCase)
                || string.Equals(current.NamespaceUri, "http://www.w3.org/2000/svg", StringComparison.Ordinal)
                || string.Equals(current.NamespaceUri, "http://www.w3.org/1998/Math/MathML", StringComparison.Ordinal)) {
                return true;
            }

            current = current.ParentElement;
        }

        return false;
    }

    private static bool IsBooleanValue(string name, string? value) {
        return string.IsNullOrEmpty(value)
            || string.Equals(value, name, StringComparison.OrdinalIgnoreCase);
    }

    private static int AttributeOrder(string name) {
        switch (name.ToLowerInvariant()) {
            case "id":
                return 0;
            case "class":
                return 1;
            case "href":
            case "src":
            case "srcset":
                return 2;
            case "alt":
            case "title":
            case "aria-label":
                return 3;
            case "style":
                return 9;
            default:
                return 5;
        }
    }


}
