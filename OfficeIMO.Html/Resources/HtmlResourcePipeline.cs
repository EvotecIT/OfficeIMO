using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Shared resource discovery and policy planning for OfficeIMO HTML workflows.
/// </summary>
public static partial class HtmlResourcePipeline {
    private const int MaxSrcDocDepth = 8;
    private const int MaxCustomPropertyResolutionDepth = 8;
    private const string ResourceSelector = "image, meta[http-equiv], [src], [srcset], [href], [xlink\\:href], [data], [data-src], [data-original], [data-original-src], [data-lazy-src], [data-srcset], [data-original-srcset], [data-lazy-srcset], [poster], [data-poster], [action], [formaction], [background], [cite], [srcdoc], [imagesrcset]";
    private const string CssCustomPropertyNamePattern = "--(?:\\\\[0-9A-Fa-f]{1,6}\\s?|\\\\[^\\r\\n\\f]|[\\p{L}\\p{N}_-]|[^\\x00-\\x7F])+";
    private static readonly Regex CssUrlExpression = new Regex("(?<name>(?:[uU]|\\\\0{0,4}(?:75|55)\\s?|\\\\[uU])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR])(?:[lL]|\\\\0{0,4}(?:6[cC]|4[cC])\\s?|\\\\[lL]))\\(\\s*(?:\"(?<url>[^\"]+)\"|'(?<url>[^']+)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssVarExpression = new Regex("(?<nameToken>(?:[vV]|\\\\0{0,4}(?:76|56)\\s?|\\\\[vV])(?:[aA]|\\\\0{0,4}(?:61|41)\\s?|\\\\[aA])(?:[rR]|\\\\0{0,4}(?:72|52)\\s?|\\\\[rR]))\\(\\s*(?<name>" + CssCustomPropertyNamePattern + ")", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssCustomPropertyDeclarationExpression = new Regex("(?<name>" + CssCustomPropertyNamePattern + ")\\s*:", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex MediaLengthFeatureExpression = new Regex("\\(\\s*(?<name>max-width|max-height|width|height)\\s*:\\s*(?<value>[+-]?(?:\\d+\\.?\\d*|\\.\\d+))\\s*(?<unit>px|em|rem|vw|vh|vmin|vmax|cm|mm|in|pt|pc)?\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    /// <summary>
    /// Parses raw HTML and builds a resource manifest.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(string html, HtmlResourcePipelineOptions? options = null) {
        IHtmlDocument document = HtmlDocumentParser.ParseDocument(html);
        return BuildManifest(document, options);
    }

    /// <summary>
    /// Builds a resource manifest from a parsed document.
    /// </summary>
    public static HtmlResourceManifest BuildManifest(IHtmlDocument document, HtmlResourcePipelineOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options = options ?? new HtmlResourcePipelineOptions();
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        var manifest = new HtmlResourceManifest();
        foreach (IElement element in document.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, element, baseUri, options, 0);
        }

        AddCssResources(manifest, document, baseUri, options);
        return manifest;
    }

    private static void AddElementResources(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options, int srcDocDepth) {
        string name = element.TagName.ToLowerInvariant();
        switch (name) {
            case "body":
            case "table":
            case "thead":
            case "tbody":
            case "tfoot":
            case "tr":
            case "td":
            case "th":
                AddLegacyBackground(manifest, element, baseUri, options);
                break;
        }

        AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "cite", baseUri, options);

        switch (name) {
            case "img":
                AddImage(manifest, element, baseUri, options);
                break;
            case "image":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "href", baseUri, options, skipFragmentOnly: true);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "xlink:href", baseUri, options, skipFragmentOnly: true);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
                break;
            case "feimage":
            case "use":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "href", baseUri, options, skipFragmentOnly: true);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "xlink:href", baseUri, options, skipFragmentOnly: true);
                break;
            case "source":
                AddSource(manifest, element, baseUri, options);
                break;
            case "link":
                AddLink(manifest, element, baseUri, options);
                break;
            case "base":
                break;
            case "meta":
                AddMetaRefresh(manifest, element, baseUri, options);
                break;
            case "a":
            case "area":
                AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "href", baseUri, options);
                break;
            case "form":
                AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "action", baseUri, options);
                break;
            case "input":
                AddSubmitterFormAction(manifest, element, baseUri, options);
                if (string.Equals((element.GetAttribute("type") ?? string.Empty).Trim(), "image", StringComparison.OrdinalIgnoreCase)) {
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "data-src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
                }

                break;
            case "button":
                AddSubmitterFormAction(manifest, element, baseUri, options);
                break;
            case "script":
                if (IsExecutableScriptElement(element)) {
                    AddAttribute(manifest, HtmlResourceKind.Script, element, "src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Script, element, "href", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Script, element, "xlink:href", baseUri, options);
                }

                break;
            case "video":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "poster", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "data-poster", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                break;
            case "audio":
            case "track":
                AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                break;
            case "object":
                AddAttribute(manifest, HtmlResourceKind.Other, element, "data", baseUri, options);
                break;
            case "embed":
                AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                break;
            case "iframe":
                if (!element.HasAttribute("srcdoc")) {
                    AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                }

                AddSrcDocResources(manifest, element, baseUri, options, srcDocDepth);
                break;
        }
    }

    private static void AddLegacyBackground(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        AddAttribute(manifest, HtmlResourceKind.Image, element, "background", baseUri, options);
    }

    private static void AddSubmitterFormAction(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!element.HasAttribute("formaction")) {
            return;
        }

        string name = element.TagName.ToLowerInvariant();
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        bool isSubmitter = string.Equals(name, "button", StringComparison.OrdinalIgnoreCase)
            ? !string.Equals(type, "button", StringComparison.OrdinalIgnoreCase) && !string.Equals(type, "reset", StringComparison.OrdinalIgnoreCase)
            : string.Equals(type, "submit", StringComparison.OrdinalIgnoreCase) || string.Equals(type, "image", StringComparison.OrdinalIgnoreCase);
        if (isSubmitter) {
            AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "formaction", baseUri, options);
        }
    }

    private static void AddImage(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (HasSelectedPictureSourceBeforeFallback(element, baseUri, options)) {
            return;
        }

        foreach (string attribute in new[] { "data-src", "data-original", "data-original-src", "data-lazy-src", "src" }) {
            AddAttribute(manifest, HtmlResourceKind.Image, element, attribute, baseUri, options);
        }

        AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
        AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-srcset", baseUri, options);
        AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-original-srcset", baseUri, options);
        AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-lazy-srcset", baseUri, options);
    }

    private static bool HasSelectedPictureSourceBeforeFallback(IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        IElement? parent = element.ParentElement;
        if (parent == null || !string.Equals(parent.TagName, "picture", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        foreach (IElement sibling in parent.Children) {
            if (ReferenceEquals(sibling, element)) {
                return false;
            }

            if (string.Equals(sibling.TagName, "source", StringComparison.OrdinalIgnoreCase)
                && HasPictureSourceCandidate(sibling)
                && HasAllowedPictureSourceCandidate(sibling, baseUri, options)
                && IsApplicableMedia(sibling.GetAttribute("media") ?? string.Empty, options.MediaContext)
                && IsSupportedPictureSourceType(sibling.GetAttribute("type"))) {
                return true;
            }
        }

        return false;
    }

    private static void AddSource(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string parentName = element.ParentElement?.TagName.ToLowerInvariant() ?? string.Empty;
        switch (parentName) {
            case "picture":
                if (IsFirstApplicablePictureSource(element, baseUri, options)
                    && HasPictureSourceCandidate(element)
                    && IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, options.MediaContext)
                    && IsSupportedPictureSourceType(element.GetAttribute("type"))) {
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "data-src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "data-original", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "data-original-src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Image, element, "data-lazy-src", baseUri, options);
                    AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
                    AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-srcset", baseUri, options);
                    AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-original-srcset", baseUri, options);
                    AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-lazy-srcset", baseUri, options);
                }

                break;
            case "audio":
            case "video":
                if (IsSelectableMediaSource(element, baseUri, options)) {
                    AddAttribute(manifest, HtmlResourceKind.Media, element, "src", baseUri, options);
                    AddAttribute(manifest, HtmlResourceKind.Media, element, "data-src", baseUri, options);
                }

                break;
        }
    }

    private static bool IsFirstApplicablePictureSource(IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        IElement? parent = element.ParentElement;
        if (parent == null || !string.Equals(parent.TagName, "picture", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        foreach (IElement sibling in parent.Children) {
            if (ReferenceEquals(sibling, element)) {
                return true;
            }

            if (!string.Equals(sibling.TagName, "source", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (HasPictureSourceCandidate(sibling)
                && HasAllowedPictureSourceCandidate(sibling, baseUri, options)
                && IsApplicableMedia(sibling.GetAttribute("media") ?? string.Empty, options.MediaContext)
                && IsSupportedPictureSourceType(sibling.GetAttribute("type"))) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSelectableMediaSource(IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        IElement? parent = element.ParentElement;
        if (parent == null) {
            return true;
        }

        string parentName = parent.TagName.ToLowerInvariant();
        if (!string.Equals(parentName, "audio", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(parentName, "video", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        if (HasNonEmptyAttribute(parent, "src") && HasAllowedMediaSourceCandidate(parent, baseUri, options)) {
            return false;
        }

        if (!IsSupportedMediaSourceType(element.GetAttribute("type"), parentName)) {
            return false;
        }

        foreach (IElement sibling in parent.Children) {
            if (ReferenceEquals(sibling, element)) {
                return true;
            }

            if (string.Equals(sibling.TagName, "source", StringComparison.OrdinalIgnoreCase)
                && HasNonEmptyAttribute(sibling, "src")
                && HasAllowedMediaSourceCandidate(sibling, baseUri, options)
                && IsSupportedMediaSourceType(sibling.GetAttribute("type"), parentName)) {
                return false;
            }
        }

        return true;
    }

    private static bool IsSupportedMediaSourceType(string? type, string parentName) {
        if (string.IsNullOrWhiteSpace(type)) {
            return true;
        }

        string mediaType = type!.Split(';')[0].Trim().ToLowerInvariant();
        if (parentName == "video") {
            return mediaType == "video/mp4" || mediaType == "video/webm" || mediaType == "video/ogg";
        }

        return mediaType == "audio/mpeg"
            || mediaType == "audio/mp4"
            || mediaType == "audio/ogg"
            || mediaType == "audio/webm"
            || mediaType == "audio/wav"
            || mediaType == "audio/wave"
            || mediaType == "audio/aac"
            || mediaType == "audio/flac";
    }

    private static bool HasNonEmptyAttribute(IElement element, string attributeName) {
        return !string.IsNullOrWhiteSpace(element.GetAttribute(attributeName));
    }

    private static bool HasPictureSourceCandidate(IElement element) {
        return HasNonEmptyAttribute(element, "srcset")
            || HasNonEmptyAttribute(element, "data-srcset")
            || HasNonEmptyAttribute(element, "data-original-srcset")
            || HasNonEmptyAttribute(element, "data-lazy-srcset")
            || HasNonEmptyAttribute(element, "src")
            || HasNonEmptyAttribute(element, "data-src")
            || HasNonEmptyAttribute(element, "data-original")
            || HasNonEmptyAttribute(element, "data-original-src")
            || HasNonEmptyAttribute(element, "data-lazy-src");
    }

    private static bool HasAllowedPictureSourceCandidate(IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(options.UrlPolicy);
        foreach (string attribute in new[] { "srcset", "data-srcset", "data-original-srcset", "data-lazy-srcset" }) {
            foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Enumerate(element.GetAttribute(attribute))) {
                if (IsAllowedResourceCandidate(HtmlResourceKind.Image, candidate.Url, baseUri, resourcePolicy)) {
                    return true;
                }
            }
        }

        foreach (string attribute in new[] { "src", "data-src", "data-original", "data-original-src", "data-lazy-src" }) {
            if (IsAllowedResourceCandidate(HtmlResourceKind.Image, element.GetAttribute(attribute), baseUri, resourcePolicy)) {
                return true;
            }
        }

        return false;
    }

    private static bool HasAllowedMediaSourceCandidate(IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        foreach (string attribute in new[] { "src", "data-src" }) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(element.GetAttribute(attribute), baseUri, options.UrlPolicy);
            if (!string.IsNullOrWhiteSpace(resolved)) {
                return true;
            }
        }

        return false;
    }

    private static void AddLink(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string rel = element.GetAttribute("rel") ?? string.Empty;
        HashSet<string> relTokens = GetRelTokens(rel);
        bool isPreload = relTokens.Contains("preload");
        bool isStylesheet = relTokens.Contains("stylesheet");
        if ((isPreload || isStylesheet) && !IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, options.MediaContext)) {
            return;
        }

        HtmlResourceKind kind;
        if (isStylesheet) {
            kind = HtmlResourceKind.Stylesheet;
        } else if (relTokens.Contains("modulepreload")) {
            kind = HtmlResourceKind.Script;
        } else if (isPreload) {
            kind = GetPreloadKind(element.GetAttribute("as"));
        } else if (relTokens.Contains("font")) {
            kind = HtmlResourceKind.Font;
        } else if (relTokens.Contains("icon") || relTokens.Contains("apple-touch-icon") || relTokens.Contains("shortcut icon")) {
            kind = HtmlResourceKind.Image;
        } else {
            kind = HtmlResourceKind.Hyperlink;
        }

        AddAttribute(manifest, kind, element, "href", baseUri, options);
        if (isPreload && kind == HtmlResourceKind.Image) {
            AddSrcSet(manifest, HtmlResourceKind.Image, element, "imagesrcset", baseUri, options);
        }
    }

    private static HashSet<string> GetRelTokens(string rel) {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string token in rel.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            string normalized = token.Trim();
            if (normalized.Length > 0) {
                tokens.Add(normalized);
            }
        }

        if (tokens.Contains("shortcut") && tokens.Contains("icon")) {
            tokens.Add("shortcut icon");
        }

        return tokens;
    }

    private static bool IsApplicableMedia(string mediaText, HtmlCssMediaContext mediaContext = HtmlCssMediaContext.Screen) {
        return HtmlComputedStyleEngine.IsApplicableMedia(mediaText, mediaContext);
    }

    private static bool ContainsMediaType(string mediaQuery, string mediaType) {
        foreach (string token in mediaQuery.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
            if (string.Equals(token.Trim(), mediaType, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool ContainsExplicitMediaType(string mediaQuery) {
        return ContainsMediaType(mediaQuery, "all")
            || ContainsMediaType(mediaQuery, "screen")
            || ContainsMediaType(mediaQuery, "print")
            || ContainsMediaType(mediaQuery, "speech");
    }

    private static IEnumerable<string> SplitTopLevelList(string text) {
        int start = 0;
        int depth = 0;
        char quote = '\0';
        for (int i = 0; i < text.Length; i++) {
            char current = text[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(text, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth = Math.Max(0, depth - 1);
                continue;
            }

            if (depth == 0 && current == ',') {
                yield return text.Substring(start, i - start);
                start = i + 1;
            }
        }

        yield return text.Substring(start);
    }

    private static HtmlResourceKind GetPreloadKind(string? asAttribute) {
        switch ((asAttribute ?? string.Empty).Trim().ToLowerInvariant()) {
            case "script":
            case "worker":
            case "serviceworker":
                return HtmlResourceKind.Script;
            case "style":
                return HtmlResourceKind.Stylesheet;
            case "image":
                return HtmlResourceKind.Image;
            case "font":
                return HtmlResourceKind.Font;
            case "audio":
            case "track":
            case "video":
                return HtmlResourceKind.Media;
            default:
                return HtmlResourceKind.Other;
        }
    }

    private static void AddCssResources(HtmlResourceManifest manifest, IHtmlDocument document, Uri? baseUri, HtmlResourcePipelineOptions options) {
        Dictionary<string, List<CssCustomPropertyDefinition>> documentCustomPropertyDefinitions = ExtractDocumentCustomPropertyDefinitions(document, options.MediaContext);
        Dictionary<IElement, int> inlineSourceOrders = GetInlineStyleSourceOrders(document, GetDocumentCssSourceOrder(document));
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, options.MediaContext)) {
                continue;
            }

            AddCssReferences(manifest, styleElement, "css", styleElement.TextContent, documentCustomPropertyDefinitions, inlineSourceOrders, sourceOrderBase: 0, includeLocalDefinitions: false, baseUri, options, document);
        }

        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            int sourceOrderBase = inlineSourceOrders.TryGetValue(element, out int inlineSourceOrder)
                ? inlineSourceOrder
                : GetDocumentCssSourceOrder(document);
            Dictionary<string, List<CssCustomPropertyDefinition>> inheritedDefinitions = ExtractInlineCustomPropertyDefinitions(element, inlineSourceOrders, options.MediaContext, includeSelf: false);
            Dictionary<string, List<CssCustomPropertyDefinition>> ambientDefinitions = MergeCustomPropertyDefinitions(documentCustomPropertyDefinitions, inheritedDefinitions);
            AddCssReferences(manifest, element, "style", element.GetAttribute("style") ?? string.Empty, ambientDefinitions, inlineSourceOrders, sourceOrderBase, includeLocalDefinitions: true, baseUri, options, document);
        }
    }

    private static bool IsCssStyleElement(IElement styleElement) {
        string type = (styleElement.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSupportedPictureSourceType(string? type) {
        return HtmlPictureSourceSupport.IsSupportedConversionContentType(type);
    }

    private static void AddSrcDocResources(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options, int srcDocDepth) {
        string? srcdoc = element.GetAttribute("srcdoc");
        if (string.IsNullOrWhiteSpace(srcdoc)) {
            return;
        }

        if (srcDocDepth >= MaxSrcDocDepth) {
            return;
        }

        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(srcdoc!);
        Uri? nestedBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, baseUri);
        foreach (IElement nestedElement in nested.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, nestedElement, nestedBaseUri, options, srcDocDepth + 1);
        }

        AddCssResources(manifest, nested, nestedBaseUri, options);
    }

    private static void AddCssReferences(
        HtmlResourceManifest manifest,
        IElement element,
        string attributeName,
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> ambientCustomPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        int sourceOrderBase,
        bool includeLocalDefinitions,
        Uri? baseUri,
        HtmlResourcePipelineOptions options,
        IHtmlDocument? document) {
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        css = StripCssCommentsOutsideStrings(css);
        List<SourceRange> inactiveMediaRanges = GetInactiveCssRuleRanges(css, options.MediaContext);
        bool scanImports = !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase);
        IElement? inlineUseElement = string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)
            ? element
            : null;
        Dictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions = includeLocalDefinitions
            ? MergeCustomPropertyDefinitions(ambientCustomPropertyDefinitions, ExtractCustomPropertyDefinitions(css, inactiveMediaRanges, sourceOrderBase, isInline: string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase), inlineOwner: inlineUseElement))
            : CloneCustomPropertyDefinitions(ambientCustomPropertyDefinitions);
        var importRanges = new List<SourceRange>();
        if (scanImports) {
            foreach (CssImportReference reference in ExtractCssImports(css)) {
                string source = reference.Source;
                if (!string.IsNullOrWhiteSpace(source)
                    && !IsInRanges(reference.Start, inactiveMediaRanges)
                    && IsApplicableCssImport(reference.ConditionText, options.MediaContext)) {
                    importRanges.Add(new SourceRange(reference.Start, reference.End));
                    AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", DecodeCssEscapes(source), baseUri, options);
                }
            }
        }

        AddUsedCustomPropertyUrls(manifest, element, attributeName, css, customPropertyDefinitions, inlineSourceOrders, inactiveMediaRanges, baseUri, options, document, inlineUseElement);
        foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
            string source = DecodeCssEscapes(reference.Source);
            if (!IsInRanges(reference.Start, inactiveMediaRanges)
                && !IsFragmentOnlyReference(source)
                && !TryGetCustomPropertyName(css, reference.Start, out _)
                && IsSupportedCssUrlDeclaration(css, reference.Start)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, reference.Start)) {
                AddRaw(manifest, ClassifyCssUrl(css, reference.Start), element, attributeName + "-image-set", source, baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            string source = DecodeCssEscapes(match.Groups["url"].Value.Trim().Trim('\'', '"'));
            if (!string.IsNullOrWhiteSpace(source)
                && !IsFragmentOnlyReference(source)
                && IsCssFunctionNameAt(css, match.Index, "url")
                && !IsImportUrl(match.Index, importRanges)
                && !IsResolvedVarFallbackUrl(css, match.Index, customPropertyDefinitions, inlineSourceOrders, document, inlineUseElement, inactiveMediaRanges, options, attributeName)
                && !IsInRanges(match.Index, inactiveMediaRanges)
                && !IsImportAtRuleUrl(css, match.Index)
                && !IsAtRulePreludeUrl(css, match.Index)
                && !IsInsideCssString(css, match.Index)
                && !IsCustomPropertyUrl(css, match.Index)
                && IsSupportedCssUrlDeclaration(css, match.Index)
                && IsCssReferenceForMatchingSelector(document, attributeName, css, match.Index)) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", source, baseUri, options);
            }
        }
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractDocumentCustomPropertyDefinitions(IHtmlDocument document, HtmlCssMediaContext mediaContext) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        int sourceOrderBase = 0;
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            string css = styleElement.TextContent;
            if (!IsCssStyleElement(styleElement) || !IsApplicableMedia(styleElement.GetAttribute("media") ?? string.Empty, mediaContext) || string.IsNullOrWhiteSpace(css)) {
                sourceOrderBase += css.Length + 1;
                continue;
            }

            css = StripCssCommentsOutsideStrings(css);
            MergeCustomPropertyDefinitionsInto(definitions, ExtractCustomPropertyDefinitions(css, GetInactiveCssRuleRanges(css, mediaContext), sourceOrderBase, isInline: false, inlineOwner: null));
            sourceOrderBase += css.Length + 1;
        }

        return definitions;
    }

    private static int GetDocumentCssSourceOrder(IHtmlDocument document) {
        int sourceOrder = 0;
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            sourceOrder += styleElement.TextContent.Length + 1;
        }

        return sourceOrder;
    }

    private static Dictionary<IElement, int> GetInlineStyleSourceOrders(IHtmlDocument document, int sourceOrderBase) {
        var sourceOrders = new Dictionary<IElement, int>();
        int sourceOrder = sourceOrderBase;
        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            sourceOrders[element] = sourceOrder;
            sourceOrder += (element.GetAttribute("style") ?? string.Empty).Length + 1;
        }

        return sourceOrders;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractInlineCustomPropertyDefinitions(IElement element, IReadOnlyDictionary<IElement, int> inlineSourceOrders, HtmlCssMediaContext mediaContext, bool includeSelf) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        for (IElement? current = includeSelf ? element : element.ParentElement; current != null; current = current.ParentElement) {
            string style = current.GetAttribute("style") ?? string.Empty;
            if (style.Length == 0 || !inlineSourceOrders.TryGetValue(current, out int sourceOrderBase)) {
                continue;
            }

            string css = StripCssCommentsOutsideStrings(style);
            MergeCustomPropertyDefinitionsInto(definitions, ExtractCustomPropertyDefinitions(css, GetInactiveCssRuleRanges(css, mediaContext), sourceOrderBase, isInline: true, inlineOwner: current));
        }

        return definitions;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> ExtractCustomPropertyDefinitions(string css, IReadOnlyList<SourceRange> inactiveMediaRanges, int sourceOrderBase, bool isInline, IElement? inlineOwner) {
        var definitions = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        foreach (Match match in CssCustomPropertyDeclarationExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            int declarationStart = match.Index;
            int valueStart = css.IndexOf(':', declarationStart);
            if (IsInsideCssString(css, declarationStart)
                || IsInRanges(declarationStart, inactiveMediaRanges)
                || valueStart < 0
                || GetCssDeclarationPropertyName(css, valueStart + 1) != propertyName) {
                continue;
            }

            int valueEnd = FindDeclarationValueEnd(css, valueStart + 1);
            string selector = GetDeclarationSelector(css, declarationStart);
            bool isImportant = IsImportantDeclarationValue(css, valueStart + 1, valueEnd);
            string valueText = GetCustomPropertyValueText(css, valueStart + 1, valueEnd);
            List<string> aliases = ExtractCustomPropertyAliases(css, valueStart + 1, valueEnd);
            bool addedUrl = false;
            foreach (Match urlMatch in CssUrlExpression.Matches(css)) {
                if (urlMatch.Index < valueStart || urlMatch.Index >= valueEnd || !IsCssFunctionNameAt(css, urlMatch.Index, "url") || IsInsideCssString(css, urlMatch.Index)) {
                    continue;
                }

                string? fallbackAlias = TryGetVarFallbackAlias(css, valueStart + 1, valueEnd, urlMatch.Index);
                AddCustomPropertyDefinition(definitions, propertyName, DecodeCssEscapes(urlMatch.Groups["url"].Value.Trim().Trim('\'', '"')), selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias);
                addedUrl = true;
            }

            foreach (CssStringUrlReference reference in ExtractImageSetStringUrls(css)) {
                if (reference.Start < valueStart || reference.Start >= valueEnd) {
                    continue;
                }

                string? fallbackAlias = TryGetVarFallbackAlias(css, valueStart + 1, valueEnd, reference.Start);
                AddCustomPropertyDefinition(definitions, propertyName, DecodeCssEscapes(reference.Source), selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias);
                addedUrl = true;
            }

            if (!addedUrl) {
                AddCustomPropertyDefinition(definitions, propertyName, string.Empty, selector, sourceOrderBase + declarationStart, isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias: null);
            }
        }

        return definitions;
    }

    private static List<string> ExtractCustomPropertyAliases(string css, int valueStart, int valueEnd) {
        var aliases = new List<string>();
        foreach (Match varMatch in CssVarExpression.Matches(css)) {
            if (varMatch.Index < valueStart
                || varMatch.Index >= valueEnd
                || !IsCssFunctionNameAt(css, varMatch.Index, "var")
                || IsInsideCssString(css, varMatch.Index)) {
                continue;
            }

            string alias = DecodeCssEscapes(varMatch.Groups["name"].Value);
            if (!aliases.Contains(alias, StringComparer.Ordinal)) {
                aliases.Add(alias);
            }
        }

        return aliases;
    }

    private static string GetCustomPropertyValueText(string css, int valueStart, int valueEnd) {
        string value = css.Substring(valueStart, Math.Max(0, valueEnd - valueStart)).Trim();
        int important = value.LastIndexOf("!important", StringComparison.OrdinalIgnoreCase);
        if (important >= 0 && string.IsNullOrWhiteSpace(value.Substring(important + 10))) {
            value = value.Substring(0, important).TrimEnd();
        }

        return DecodeCssEscapes(value).Trim();
    }

    private static string? TryGetVarFallbackAlias(string css, int valueStart, int valueEnd, int urlIndex) {
        foreach (Match varMatch in CssVarExpression.Matches(css)) {
            if (varMatch.Index < valueStart
                || varMatch.Index >= valueEnd
                || urlIndex <= varMatch.Index
                || !IsCssFunctionNameAt(css, varMatch.Index, "var")
                || IsInsideCssString(css, varMatch.Index)) {
                continue;
            }

            int open = css.IndexOf('(', varMatch.Index);
            if (open < 0) {
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close < 0 || close > valueEnd || urlIndex >= close) {
                continue;
            }

            int comma = FindTopLevelComma(css, open + 1, close);
            if (comma >= 0 && urlIndex > comma) {
                return DecodeCssEscapes(varMatch.Groups["name"].Value);
            }
        }

        return null;
    }

    private static bool IsImportantDeclarationValue(string css, int valueStart, int valueEnd) {
        int index = valueEnd - 1;
        while (index >= valueStart && char.IsWhiteSpace(css[index])) {
            index--;
        }

        const string Important = "important";
        if (index - Important.Length + 1 < valueStart) {
            return false;
        }

        string suffix = css.Substring(index - Important.Length + 1, Important.Length);
        if (!string.Equals(suffix, Important, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        index -= Important.Length;
        while (index >= valueStart && char.IsWhiteSpace(css[index])) {
            index--;
        }

        return index >= valueStart && css[index] == '!';
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> CloneCustomPropertyDefinitions(IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> definitions) {
        var clone = new Dictionary<string, List<CssCustomPropertyDefinition>>(StringComparer.Ordinal);
        MergeCustomPropertyDefinitionsInto(clone, definitions);
        return clone;
    }

    private static Dictionary<string, List<CssCustomPropertyDefinition>> MergeCustomPropertyDefinitions(
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> first,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> second) {
        Dictionary<string, List<CssCustomPropertyDefinition>> merged = CloneCustomPropertyDefinitions(first);
        MergeCustomPropertyDefinitionsInto(merged, second);
        return merged;
    }

    private static void MergeCustomPropertyDefinitionsInto(
        IDictionary<string, List<CssCustomPropertyDefinition>> target,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> source) {
        foreach (KeyValuePair<string, List<CssCustomPropertyDefinition>> pair in source) {
            if (!target.TryGetValue(pair.Key, out List<CssCustomPropertyDefinition>? values)) {
                values = new List<CssCustomPropertyDefinition>();
                target[pair.Key] = values;
            }

            values.AddRange(pair.Value);
        }
    }

    private static List<SourceRange> GetInactiveCssRuleRanges(string css, HtmlCssMediaContext mediaContext) {
        List<SourceRange> ranges = GetInactiveMediaRanges(css, mediaContext);
        ranges.AddRange(GetInactiveSupportsRanges(css));
        return ranges;
    }

    private static List<SourceRange> GetInactiveMediaRanges(string css, HtmlCssMediaContext mediaContext) {
        var ranges = new List<SourceRange>();
        int index = 0;
        while (index < css.Length) {
            int mediaStart = css.IndexOf("@media", index, StringComparison.OrdinalIgnoreCase);
            if (mediaStart < 0) {
                break;
            }

            if (IsInsideCssString(css, mediaStart) || !HasAtRuleTokenBoundary(css, mediaStart, "@media")) {
                index = mediaStart + 6;
                continue;
            }

            int preludeStart = mediaStart + 6;
            int open = FindNextTopLevelBlockStart(css, preludeStart);
            if (open < 0) {
                break;
            }

            int close = FindMatchingCssBrace(css, open);
            if (close <= open) {
                break;
            }

            string mediaText = css.Substring(preludeStart, open - preludeStart).Trim();
            if (!IsApplicableMedia(mediaText, mediaContext)) {
                ranges.Add(new SourceRange(open + 1, close));
                index = close + 1;
            } else {
                index = open + 1;
            }
        }

        return ranges;
    }

    private static List<SourceRange> GetInactiveSupportsRanges(string css) {
        var ranges = new List<SourceRange>();
        int index = 0;
        while (index < css.Length) {
            int supportsStart = css.IndexOf("@supports", index, StringComparison.OrdinalIgnoreCase);
            if (supportsStart < 0) {
                break;
            }

            if (IsInsideCssString(css, supportsStart) || !HasAtRuleTokenBoundary(css, supportsStart, "@supports")) {
                index = supportsStart + 9;
                continue;
            }

            int preludeStart = supportsStart + 9;
            int open = FindNextTopLevelBlockStart(css, preludeStart);
            if (open < 0) {
                break;
            }

            int close = FindMatchingCssBrace(css, open);
            if (close <= open) {
                break;
            }

            string conditionText = css.Substring(preludeStart, open - preludeStart).Trim();
            if (!HtmlComputedStyleEngine.IsApplicableSupports(conditionText)) {
                ranges.Add(new SourceRange(open + 1, close));
                index = close + 1;
            } else {
                index = open + 1;
            }
        }

        return ranges;
    }

    private static int FindNextTopLevelBlockStart(string css, int start) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth = Math.Max(0, depth - 1);
                continue;
            }

            if (depth == 0) {
                if (current == '{') {
                    return i;
                }

                if (current == ';') {
                    return -1;
                }
            }
        }

        return -1;
    }

    private static int FindMatchingCssBrace(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '{') {
                depth++;
                continue;
            }

            if (current == '}') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static void AddCustomPropertyDefinition(IDictionary<string, List<CssCustomPropertyDefinition>> definitions, string propertyName, string source, string selector, int declarationStart, bool isImportant, IReadOnlyList<string> aliases, bool isInline, IElement? inlineOwner, string valueText, string? fallbackAlias) {
        if (!definitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? values)) {
            values = new List<CssCustomPropertyDefinition>();
            definitions[propertyName] = values;
        }

        values.Add(new CssCustomPropertyDefinition(source, selector, declarationStart, !string.IsNullOrWhiteSpace(source), isImportant, aliases, isInline, inlineOwner, valueText, fallbackAlias));
    }

    private static void AddUsedCustomPropertyUrls(
        HtmlResourceManifest manifest,
        IElement element,
        string attributeName,
        string css,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IReadOnlyList<SourceRange> inactiveRanges,
        Uri? baseUri,
        HtmlResourcePipelineOptions options,
        IHtmlDocument? document,
        IElement? useElement) {
        if (customPropertyDefinitions.Count == 0) {
            return;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            if (IsInRanges(match.Index, inactiveRanges)
                || !IsCssFunctionNameAt(css, match.Index, "var")
                || IsInsideCssString(css, match.Index)) {
                continue;
            }

            HtmlResourceKind kind = ClassifyCssUrl(css, match.Index);
            if (kind == HtmlResourceKind.Other) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            if (!IsCssReferenceForMatchingSelector(document, attributeName, css, match.Index)) {
                continue;
            }

            var addedSources = new HashSet<string>(StringComparer.Ordinal);
            if (document != null && useElement == null && !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
                IElement[] matchedElements = GetElementsMatchingSelectorList(document, useSelector).ToArray();
                if (matchedElements.Length > 0) {
                    foreach (IElement matchedElement in matchedElements) {
                        Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions = ExtractInlineCustomPropertyDefinitions(matchedElement, inlineSourceOrders, options.MediaContext, includeSelf: true);
                        Dictionary<string, List<CssCustomPropertyDefinition>> mergedDefinitions = inlineDefinitions.Count == 0
                            ? CloneCustomPropertyDefinitions(customPropertyDefinitions)
                            : MergeCustomPropertyDefinitions(customPropertyDefinitions, inlineDefinitions);
                        foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(propertyName, mergedDefinitions, useSelector, document, matchedElement, new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                            if (!IsFragmentOnlyReference(source.Source) && addedSources.Add(source.Source)) {
                                AddRaw(manifest, kind, element, attributeName + "-var-url", source.Source, baseUri, options);
                            }
                        }
                    }

                    continue;
                }
            }

            foreach (CssCustomPropertyDefinition source in ResolveCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, useSelector, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0)) {
                if (!IsFragmentOnlyReference(source.Source) && addedSources.Add(source.Source)) {
                    AddRaw(manifest, kind, element, attributeName + "-var-url", source.Source, baseUri, options);
                }
            }
        }
    }

    private static IEnumerable<CssCustomPropertyDefinition> ResolveCustomPropertyUrlDefinitions(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        string useSelector,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth) {
        if (depth >= MaxCustomPropertyResolutionDepth
            || !visited.Add(propertyName)
            || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
            yield break;
        }

        int selectedDeclarationStart = SelectCustomPropertyDeclaration(sources, useSelector, document, useElement);
        if (selectedDeclarationStart < 0) {
            visited.Remove(propertyName);
            yield break;
        }

        foreach (CssCustomPropertyDefinition source in sources) {
            if (source.DeclarationStart != selectedDeclarationStart || !CanSubstituteCustomProperty(source, useSelector, document, useElement)) {
                continue;
            }

            if (source.IsInheritedKeyword) {
                foreach (CssCustomPropertyDefinition inheritedSource in ResolveInheritedCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, document, useElement, visited, depth)) {
                    yield return inheritedSource;
                }

                continue;
            }

            if (source.HasUrl) {
                if (source.FallbackAlias == null || !HasResolvedCustomProperty(source.FallbackAlias, customPropertyDefinitions, document, useElement, visited, depth + 1)) {
                    yield return source;
                }
            }

            foreach (string alias in source.Aliases) {
                foreach (CssCustomPropertyDefinition aliasSource in ResolveCustomPropertyUrlDefinitions(alias, customPropertyDefinitions, useSelector, document, useElement, visited, depth + 1)) {
                    yield return aliasSource;
                }
            }
        }

        visited.Remove(propertyName);
    }

    private static IEnumerable<CssCustomPropertyDefinition> ResolveInheritedCustomPropertyUrlDefinitions(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth) {
        if (useElement?.ParentElement == null) {
            yield break;
        }

        visited.Remove(propertyName);
        foreach (CssCustomPropertyDefinition inheritedSource in ResolveCustomPropertyUrlDefinitions(propertyName, customPropertyDefinitions, string.Empty, document, useElement.ParentElement, visited, depth + 1)) {
            yield return inheritedSource;
        }

        visited.Add(propertyName);
    }

    private static bool CanSubstituteCustomProperty(CssCustomPropertyDefinition source, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        string definitionSelector = source.Selector;
        if (string.IsNullOrWhiteSpace(definitionSelector)) {
            if (source.IsInline && useElement != null) {
                return GetInlineOwnerDistance(source, useElement) != int.MaxValue;
            }

            return string.IsNullOrWhiteSpace(useSelector);
        }

        if (string.Equals(definitionSelector, useSelector, StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (SelectorMatchesElementOrAncestor(normalizedDefinition, useElement)) {
                return true;
            }

            if (string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }

            if (useElement != null) {
                continue;
            }

            foreach (string usePart in SplitTopLevelList(useSelector)) {
                string normalizedUse = usePart.Trim();
                if (IsAncestorSelector(normalizedDefinition, normalizedUse)
                    || SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)
                    || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                    || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                    return true;
                }
            }
        }

        return false;
    }

    private static int SelectCustomPropertyDeclaration(IEnumerable<CssCustomPropertyDefinition> sources, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        int selectedDeclarationStart = -1;
        int selectedRank = -1;
        int selectedSpecificity = -1;
        int selectedDistance = int.MaxValue;
        bool selectedImportant = false;
        foreach (CssCustomPropertyDefinition source in sources) {
            int rank = GetSubstitutionRank(source, useSelector, document, useElement);
            if (rank < 0) {
                continue;
            }

            int distance = GetElementSubstitutionDistance(source, useElement);
            int specificity = GetMatchingSelectorSpecificity(source.Selector, useSelector, document, useElement);
            bool sameElementCascade = rank >= 3 && selectedRank >= 3;
            if ((sameElementCascade && source.IsImportant != selectedImportant && source.IsImportant)
                || (!(sameElementCascade && source.IsImportant != selectedImportant) && rank > selectedRank)
                || (rank == selectedRank
                    && (distance < selectedDistance
                        || (distance == selectedDistance
                            && ((!selectedImportant && source.IsImportant)
                                || (source.IsImportant == selectedImportant
                                    && (specificity > selectedSpecificity
                                        || (specificity == selectedSpecificity && source.DeclarationStart >= selectedDeclarationStart)))))))) {
                selectedImportant = source.IsImportant;
                selectedRank = rank;
                selectedSpecificity = specificity;
                selectedDistance = distance;
                selectedDeclarationStart = source.DeclarationStart;
            }
        }

        return selectedRank >= 0 ? selectedDeclarationStart : -1;
    }

    private static int GetSubstitutionRank(CssCustomPropertyDefinition source, string useSelector, IHtmlDocument? document = null, IElement? useElement = null) {
        string definitionSelector = source.Selector;
        if (string.IsNullOrWhiteSpace(definitionSelector)) {
            if (source.IsInline && useElement != null) {
                int inlineDistance = GetInlineOwnerDistance(source, useElement);
                if (inlineDistance == int.MaxValue) {
                    return -1;
                }

                return inlineDistance == 0 ? 4 : 2;
            }

            return string.IsNullOrWhiteSpace(useSelector) ? 3 : -1;
        }

        int best = -1;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            best = Math.Max(best, GetElementSubstitutionRank(normalizedDefinition, useElement));
            if (string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase)) {
                best = Math.Max(best, 1);
            }

            if (useElement != null) {
                continue;
            }

            foreach (string usePart in SplitTopLevelList(useSelector)) {
                string normalizedUse = usePart.Trim();
                if (string.Equals(normalizedDefinition, normalizedUse, StringComparison.OrdinalIgnoreCase)) {
                    best = Math.Max(best, 3);
                } else if (SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)) {
                    best = Math.Max(best, 3);
                } else if (IsAncestorSelector(normalizedDefinition, normalizedUse)
                    || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                    || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                    best = Math.Max(best, 2);
                }
            }
        }

        return best;
    }

    private static int GetElementSubstitutionRank(string definitionSelector, IElement? useElement) {
        if (useElement == null || string.IsNullOrWhiteSpace(definitionSelector)) {
            return -1;
        }

        if (ElementMatchesSelector(useElement, definitionSelector)) {
            return 3;
        }

        for (IElement? ancestor = useElement.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (ElementMatchesSelector(ancestor, definitionSelector)) {
                return 2;
            }
        }

        return -1;
    }

    private static int GetElementSubstitutionDistance(CssCustomPropertyDefinition source, IElement? useElement) {
        if (source.IsInline) {
            return GetInlineOwnerDistance(source, useElement);
        }

        string definitionSelector = source.Selector;
        if (useElement == null || string.IsNullOrWhiteSpace(definitionSelector)) {
            return int.MaxValue;
        }

        int best = int.MaxValue;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (ElementMatchesSelector(useElement, normalizedDefinition)) {
                best = Math.Min(best, 0);
                continue;
            }

            int distance = 1;
            for (IElement? ancestor = useElement.ParentElement; ancestor != null; ancestor = ancestor.ParentElement, distance++) {
                if (ElementMatchesSelector(ancestor, normalizedDefinition)) {
                    best = Math.Min(best, distance);
                    break;
                }
            }
        }

        return best;
    }

    private static int GetInlineOwnerDistance(CssCustomPropertyDefinition source, IElement? useElement) {
        if (!source.IsInline || source.InlineOwner == null || useElement == null) {
            return int.MaxValue;
        }

        int distance = 0;
        for (IElement? current = useElement; current != null; current = current.ParentElement, distance++) {
            if (ReferenceEquals(current, source.InlineOwner)) {
                return distance;
            }
        }

        return int.MaxValue;
    }

    private static int GetMatchingSelectorSpecificity(string definitionSelector, string useSelector, IHtmlDocument? document, IElement? useElement) {
        int best = -1;
        foreach (string definitionPart in SplitTopLevelList(definitionSelector)) {
            string normalizedDefinition = definitionPart.Trim();
            if (normalizedDefinition.Length == 0) {
                continue;
            }

            bool matches = SelectorMatchesElementOrAncestor(normalizedDefinition, useElement)
                || string.Equals(normalizedDefinition, ":root", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "html", StringComparison.OrdinalIgnoreCase)
                || string.Equals(normalizedDefinition, "body", StringComparison.OrdinalIgnoreCase);
            if (!matches) {
                if (useElement != null) {
                    continue;
                }

                foreach (string usePart in SplitTopLevelList(useSelector)) {
                    string normalizedUse = usePart.Trim();
                    if (string.Equals(normalizedDefinition, normalizedUse, StringComparison.OrdinalIgnoreCase)
                        || SelectorSameElementMatches(document, normalizedDefinition, normalizedUse)
                        || IsAncestorSelector(normalizedDefinition, normalizedUse)
                        || IsSameElementSelectorPrefix(normalizedDefinition, normalizedUse)
                        || SelectorRelationshipMatches(document, normalizedDefinition, normalizedUse)) {
                        matches = true;
                        break;
                    }
                }
            }

            if (matches) {
                best = Math.Max(best, CalculateSelectorSpecificity(normalizedDefinition));
            }
        }

        return best;
    }

    private static int CalculateSelectorSpecificity(string selector) {
        int ids = 0;
        int classesAttributesAndPseudoClasses = 0;
        int elements = 0;
        bool inAttribute = false;
        for (int i = 0; i < selector.Length; i++) {
            char current = selector[i];
            if (current == '[') {
                inAttribute = true;
                classesAttributesAndPseudoClasses++;
                continue;
            }

            if (current == ']') {
                inAttribute = false;
                continue;
            }

            if (inAttribute) {
                continue;
            }

            if (current == '#') {
                ids++;
                i = SkipCssIdentifier(selector, i + 1);
            } else if (current == '.') {
                classesAttributesAndPseudoClasses++;
                i = SkipCssIdentifier(selector, i + 1);
            } else if (current == ':') {
                if (i + 1 < selector.Length && selector[i + 1] == ':') {
                    elements++;
                    i = SkipCssIdentifier(selector, i + 2);
                } else {
                    if (TryReadPseudoClassName(selector, i + 1, out string pseudoClassName, out int nameEnd)) {
                        if (nameEnd < selector.Length && selector[nameEnd] == '(') {
                            int close = FindMatchingCssParenthesis(selector, nameEnd);
                            if (close > nameEnd) {
                                string argument = selector.Substring(nameEnd + 1, close - nameEnd - 1);
                                if (string.Equals(pseudoClassName, "where", StringComparison.OrdinalIgnoreCase)) {
                                    i = close;
                                    continue;
                                }

                                if (string.Equals(pseudoClassName, "is", StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(pseudoClassName, "not", StringComparison.OrdinalIgnoreCase)
                                    || string.Equals(pseudoClassName, "has", StringComparison.OrdinalIgnoreCase)) {
                                    int argumentSpecificity = MaxSelectorSpecificity(argument);
                                    ids += argumentSpecificity / 10000;
                                    classesAttributesAndPseudoClasses += (argumentSpecificity % 10000) / 100;
                                    elements += argumentSpecificity % 100;
                                    i = close;
                                    continue;
                                }

                                classesAttributesAndPseudoClasses++;
                                i = close;
                                continue;
                            }
                        }

                        classesAttributesAndPseudoClasses++;
                        i = nameEnd - 1;
                        continue;
                    }

                    classesAttributesAndPseudoClasses++;
                    i = SkipCssIdentifier(selector, i + 1);
                }
            } else if (IsSelectorElementStart(selector, i)) {
                elements++;
                i = SkipCssIdentifier(selector, i);
            }
        }

        return (ids * 10000) + (classesAttributesAndPseudoClasses * 100) + elements;
    }

    private static int MaxSelectorSpecificity(string selectorList) {
        int max = 0;
        foreach (string selector in SplitTopLevelList(selectorList)) {
            int specificity = CalculateSelectorSpecificity(selector);
            if (specificity > max) {
                max = specificity;
            }
        }

        return max;
    }

    private static int SkipCssIdentifier(string selector, int start) {
        int cursor = start;
        while (cursor < selector.Length && (IsCssIdentifierCharacter(selector[cursor]) || selector[cursor] == '\\')) {
            cursor++;
        }

        return Math.Max(start, cursor) - 1;
    }

    private static bool IsSelectorElementStart(string selector, int index) {
        char current = selector[index];
        if (!char.IsLetter(current) && current != '*') {
            return false;
        }

        if (current == '*') {
            return false;
        }

        if (index > 0) {
            char previous = selector[index - 1];
            if (previous == '#'
                || previous == '.'
                || previous == ':'
                || previous == '-'
                || previous == '_'
                || char.IsLetterOrDigit(previous)) {
                return false;
            }
        }

        return true;
    }

    private static bool SelectorMatchesElementOrAncestor(string selector, IElement? useElement) {
        return GetElementSubstitutionRank(selector, useElement) >= 0;
    }

    private static bool ElementMatchesSelector(IElement element, string selector) {
        string normalized = NormalizeSelectorForQuery(selector, stripPseudoElements: false, stripStatefulPseudoClasses: true);
        if (normalized.Length == 0 || normalized.StartsWith("@", StringComparison.Ordinal)) {
            return false;
        }

        try {
            return element.Matches(normalized);
        } catch {
            return false;
        }
    }

    private static bool IsResolvedVarFallbackUrl(
        string css,
        int urlIndex,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument? document,
        IElement? useElement,
        IReadOnlyList<SourceRange> inactiveRanges,
        HtmlResourcePipelineOptions options,
        string attributeName) {
        if (customPropertyDefinitions.Count == 0) {
            return false;
        }

        foreach (Match match in CssVarExpression.Matches(css)) {
            string propertyName = DecodeCssEscapes(match.Groups["name"].Value);
            if (IsInRanges(match.Index, inactiveRanges)
                || !IsCssFunctionNameAt(css, match.Index, "var")
                || IsInsideCssString(css, match.Index)
                || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
                continue;
            }

            int open = css.IndexOf('(', match.Index);
            if (open < 0) {
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                continue;
            }

            int comma = FindTopLevelComma(css, open + 1, close);
            if (comma < 0 || urlIndex <= comma || urlIndex >= close) {
                continue;
            }

            string useSelector = GetDeclarationSelector(css, match.Index);
            if (document != null && useElement == null && !string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
                IElement[] matchedElements = GetElementsMatchingSelectorList(document, useSelector).ToArray();
                if (matchedElements.Length > 0) {
                    return matchedElements.All(matchedElement => HasResolvedCustomProperty(propertyName, customPropertyDefinitions, inlineSourceOrders, document, matchedElement, options, useSelector));
                }
            }

            return HasResolvedCustomProperty(propertyName, customPropertyDefinitions, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0, useSelector: useSelector);
        }

        return false;
    }

    private static bool HasResolvedCustomProperty(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IReadOnlyDictionary<IElement, int> inlineSourceOrders,
        IHtmlDocument document,
        IElement useElement,
        HtmlResourcePipelineOptions options,
        string useSelector) {
        Dictionary<string, List<CssCustomPropertyDefinition>> inlineDefinitions = ExtractInlineCustomPropertyDefinitions(useElement, inlineSourceOrders, options.MediaContext, includeSelf: true);
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> mergedDefinitions = inlineDefinitions.Count == 0
            ? customPropertyDefinitions
            : MergeCustomPropertyDefinitions(customPropertyDefinitions, inlineDefinitions);
        return HasResolvedCustomProperty(propertyName, mergedDefinitions, document, useElement, new HashSet<string>(StringComparer.Ordinal), depth: 0, useSelector: useSelector);
    }

    private static bool HasResolvedCustomProperty(
        string propertyName,
        IReadOnlyDictionary<string, List<CssCustomPropertyDefinition>> customPropertyDefinitions,
        IHtmlDocument? document,
        IElement? useElement,
        ISet<string> visited,
        int depth,
        string useSelector = "") {
        if (depth >= MaxCustomPropertyResolutionDepth
            || !visited.Add(propertyName)
            || !customPropertyDefinitions.TryGetValue(propertyName, out List<CssCustomPropertyDefinition>? sources)) {
            return false;
        }

        bool resolved = false;
        int selectedDeclarationStart = SelectCustomPropertyDeclaration(sources, useSelector, document, useElement);
        if (selectedDeclarationStart >= 0) {
            foreach (CssCustomPropertyDefinition source in sources) {
                if (source.DeclarationStart != selectedDeclarationStart || !CanSubstituteCustomProperty(source, useSelector, document, useElement)) {
                    continue;
                }

                if (source.IsInheritedKeyword) {
                    if (useElement?.ParentElement != null) {
                        visited.Remove(propertyName);
                        resolved = HasResolvedCustomProperty(propertyName, customPropertyDefinitions, document, useElement.ParentElement, visited, depth + 1);
                        visited.Add(propertyName);
                    }
                } else if (source.HasUrl) {
                    resolved = source.FallbackAlias == null
                        || !HasResolvedCustomProperty(source.FallbackAlias, customPropertyDefinitions, document, useElement, visited, depth + 1, useSelector);
                } else if (source.Aliases.Count == 0 && !source.IsCssWideInvalidatingKeyword) {
                    resolved = true;
                } else {
                    foreach (string alias in source.Aliases) {
                        if (HasResolvedCustomProperty(alias, customPropertyDefinitions, document, useElement, visited, depth + 1, useSelector)) {
                            resolved = true;
                            break;
                        }
                    }
                }

                if (resolved) {
                    break;
                }
            }
        }

        visited.Remove(propertyName);
        return resolved;
    }

    private static int FindDeclarationValueEnd(string css, int start) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth = Math.Max(0, depth - 1);
                continue;
            }

            if (depth == 0 && (current == ';' || current == '}')) {
                return i;
            }
        }

        return css.Length;
    }

    private static int FindTopLevelComma(string css, int start, int end) {
        int depth = 0;
        char quote = '\0';
        for (int i = start; i < end; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                if (depth > 0) {
                    depth--;
                }

                continue;
            }

            if (depth == 0 && current == ',') {
                return i;
            }
        }

        return -1;
    }

    private static bool IsAncestorSelector(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return char.IsWhiteSpace(next) || next == '>';
    }

    private static bool IsSameElementSelectorPrefix(string definitionSelector, string useSelector) {
        if (definitionSelector.Length == 0 || useSelector.Length <= definitionSelector.Length) {
            return false;
        }

        if (!useSelector.StartsWith(definitionSelector, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        char next = useSelector[definitionSelector.Length];
        return next == '.'
            || next == '#'
            || next == '['
            || next == ':';
    }

    private static string GetDeclarationSelector(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        if (blockStart < 0) {
            return string.Empty;
        }

        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, blockStart - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, blockStart - 1));
        int selectorStart = Math.Max(0, Math.Max(previousBlockEnd, previousStatementEnd) + 1);
        string selector = css.Substring(selectorStart, blockStart - selectorStart).Trim();
        int groupingStart = selector.LastIndexOf('{');
        return groupingStart >= 0
            ? selector.Substring(groupingStart + 1).Trim()
            : selector;
    }

    private static bool IsCssReferenceForMatchingSelector(IHtmlDocument? document, string attributeName, string css, int index) {
        if (document == null || string.Equals(attributeName, "style", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }

        string selector = GetDeclarationSelector(css, index);
        if (string.IsNullOrWhiteSpace(selector) || selector.TrimStart().StartsWith("@", StringComparison.Ordinal)) {
            return true;
        }

        foreach (string selectorPart in SplitTopLevelList(selector)) {
            string normalized = NormalizeSelectorForQuery(selectorPart, stripStatefulPseudoClasses: true);
            if (normalized.Length == 0) {
                if (IsBarePseudoElementSelector(selectorPart) || IsStatefulPseudoClassOnlySelector(selectorPart)) {
                    return true;
                }

                continue;
            }

            try {
                if (document.QuerySelector(normalized) != null) {
                    return true;
                }
            } catch {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<IElement> GetElementsMatchingSelectorList(IHtmlDocument document, string selector) {
        if (string.IsNullOrWhiteSpace(selector) || selector.TrimStart().StartsWith("@", StringComparison.Ordinal)) {
            yield break;
        }

        var seen = new HashSet<IElement>();
        foreach (string selectorPart in SplitTopLevelList(selector)) {
            string normalized = NormalizeSelectorForQuery(selectorPart, stripStatefulPseudoClasses: true);
            if (normalized.Length == 0) {
                continue;
            }

            IEnumerable<IElement> matches;
            try {
                matches = document.QuerySelectorAll(normalized).OfType<IElement>().ToArray();
            } catch {
                continue;
            }

            foreach (IElement match in matches) {
                if (seen.Add(match)) {
                    yield return match;
                }
            }
        }
    }

    private static bool SelectorRelationshipMatches(IHtmlDocument? document, string definitionSelector, string useSelector) {
        if (document == null || string.IsNullOrWhiteSpace(definitionSelector) || string.IsNullOrWhiteSpace(useSelector)) {
            return false;
        }

        string normalizedDefinition = NormalizeSelectorForQuery(definitionSelector, stripPseudoElements: false);
        string normalizedUse = NormalizeSelectorForQuery(useSelector);
        if (normalizedDefinition.Length == 0 || normalizedUse.Length == 0) {
            return false;
        }

        try {
            if (document.QuerySelector(normalizedDefinition + " " + normalizedUse) != null) {
                return true;
            }

            foreach (IElement useMatch in document.QuerySelectorAll(normalizedUse)) {
                for (IElement? ancestor = useMatch.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
                    if (ancestor.Matches(normalizedDefinition)) {
                        return true;
                    }
                }
            }

            return false;
        } catch {
            return false;
        }
    }

    private static bool SelectorSameElementMatches(IHtmlDocument? document, string definitionSelector, string useSelector) {
        if (document == null || string.IsNullOrWhiteSpace(definitionSelector) || string.IsNullOrWhiteSpace(useSelector)) {
            return false;
        }

        string normalizedDefinition = NormalizeSelectorForQuery(definitionSelector, stripPseudoElements: false);
        string normalizedUse = NormalizeSelectorForQuery(useSelector);
        if (normalizedDefinition.Length == 0 || normalizedUse.Length == 0) {
            return false;
        }

        try {
            foreach (IElement useMatch in document.QuerySelectorAll(normalizedUse)) {
                if (useMatch.Matches(normalizedDefinition)) {
                    return true;
                }
            }

            return false;
        } catch {
            return false;
        }
    }

    private static string NormalizeSelectorForQuery(string selector, bool stripPseudoElements = true, bool stripStatefulPseudoClasses = false) {
        string normalized = selector.Trim();
        int pseudoElement = stripPseudoElements ? normalized.IndexOf("::", StringComparison.Ordinal) : -1;
        if (pseudoElement >= 0) {
            normalized = normalized.Substring(0, pseudoElement).TrimEnd();
        }

        if (stripStatefulPseudoClasses) {
            normalized = StripStatefulPseudoClasses(normalized).Trim();
        }

        return normalized;
    }

    private static bool IsBarePseudoElementSelector(string selector) {
        string trimmed = selector.Trim();
        return trimmed.StartsWith("::", StringComparison.Ordinal);
    }

    private static bool IsStatefulPseudoClassOnlySelector(string selector) {
        string stripped = StripStatefulPseudoClasses(selector.Trim()).Trim();
        return stripped.Length == 0;
    }

    private static string StripStatefulPseudoClasses(string selector) {
        var result = new StringBuilder(selector.Length);
        for (int i = 0; i < selector.Length; i++) {
            if (selector[i] == ':'
                && (i + 1 >= selector.Length || selector[i + 1] != ':')
                && TryReadPseudoClassName(selector, i + 1, out string pseudoClassName, out int nameEnd)
                && IsStatefulPseudoClass(pseudoClassName)) {
                i = nameEnd - 1;
                continue;
            }

            result.Append(selector[i]);
        }

        return result.ToString();
    }

    private static bool TryReadPseudoClassName(string selector, int start, out string name, out int end) {
        int cursor = start;
        while (cursor < selector.Length && (char.IsLetterOrDigit(selector[cursor]) || selector[cursor] == '-')) {
            cursor++;
        }

        if (cursor == start) {
            name = string.Empty;
            end = start;
            return false;
        }

        name = selector.Substring(start, cursor - start);
        end = cursor;
        return true;
    }

    private static bool IsStatefulPseudoClass(string pseudoClassName) {
        switch (pseudoClassName.ToLowerInvariant()) {
            case "active":
            case "focus":
            case "focus-visible":
            case "focus-within":
            case "hover":
            case "target":
            case "visited":
                return true;
            default:
                return false;
        }
    }

    private static int GetDeclarationStart(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousStatementEnd = css.LastIndexOf(';', Math.Max(0, index - 1));
        return Math.Max(0, Math.Max(blockStart, previousStatementEnd) + 1);
    }

    private static IEnumerable<CssStringUrlReference> ExtractImageSetStringUrls(string css) {
        int index = 0;
        while (index < css.Length) {
            if (!TryFindNextCssFunction(css, index, out int functionStart, out int open, "image-set", "-webkit-image-set")) {
                yield break;
            }

            if (IsInsideCssString(css, functionStart)) {
                index = open + 1;
                continue;
            }

            int close = FindMatchingCssParenthesis(css, open);
            if (close <= open) {
                yield break;
            }

            int valueCursor = open + 1;
            while (valueCursor < close) {
                char current = css[valueCursor];
                if ((current == '"' || current == '\'') && !IsCssTypeFunctionString(css, valueCursor)) {
                    if (TryReadCssQuotedValue(css, valueCursor, out string source, out int end)) {
                        if (!string.IsNullOrWhiteSpace(source)) {
                            yield return new CssStringUrlReference(functionStart, end, source);
                        }

                        valueCursor = end;
                        continue;
                    }
                }

                valueCursor++;
            }

            index = close + 1;
        }
    }

    private static bool IsCssTypeFunctionString(string css, int quoteIndex) {
        int cursor = quoteIndex - 1;
        cursor = SkipCssWhitespaceAndCommentsBackward(css, cursor);

        if (cursor < 0 || css[cursor] != '(') {
            return false;
        }

        cursor--;
        cursor = SkipCssWhitespaceAndCommentsBackward(css, cursor);

        int end = cursor + 1;
        while (cursor >= 0 && (IsCssIdentifierCharacter(css[cursor]) || css[cursor] == '\\')) {
            cursor--;
        }

        string functionName = css.Substring(cursor + 1, end - cursor - 1);
        return CssFunctionNameEquals(functionName, "type");
    }

    private static int SkipCssWhitespaceAndCommentsBackward(string css, int cursor) {
        while (cursor >= 0) {
            if (char.IsWhiteSpace(css[cursor])) {
                cursor--;
                continue;
            }

            if (cursor > 0 && css[cursor - 1] == '*' && css[cursor] == '/') {
                int commentStart = css.LastIndexOf("/*", cursor - 2, StringComparison.Ordinal);
                if (commentStart < 0) {
                    return cursor;
                }

                cursor = commentStart - 1;
                continue;
            }

            break;
        }

        return cursor;
    }

    private static int FindMatchingCssParenthesis(string css, int open) {
        int depth = 0;
        char quote = '\0';
        for (int i = open; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '(') {
                depth++;
                continue;
            }

            if (current == ')') {
                depth--;
                if (depth == 0) {
                    return i;
                }
            }
        }

        return -1;
    }

    private static IEnumerable<CssImportReference> ExtractCssImports(string css) {
        int index = 0;
        while (index < css.Length) {
            int importStart = css.IndexOf("@import", index, StringComparison.OrdinalIgnoreCase);
            if (importStart < 0) {
                yield break;
            }

            if (IsInsideCssString(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (!HasImportTokenBoundary(css, importStart)) {
                index = importStart + 7;
                continue;
            }

            if (HasStyleRuleBefore(css, importStart)) {
                yield break;
            }

            int cursor = SkipWhitespace(css, importStart + 7);
            string source;
            int end;
            if (IsCssFunctionNameAt(css, cursor, "url")) {
                int open = css.IndexOf('(', cursor);
                cursor = SkipWhitespace(css, open + 1);
                if (!TryReadCssUrlFunctionSource(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
                if (!TryReadCssQuotedValue(css, cursor, out source, out end)) {
                    index = importStart + 7;
                    continue;
                }
            } else {
                int sourceStart = cursor;
                while (cursor < css.Length && !char.IsWhiteSpace(css[cursor]) && css[cursor] != ';') {
                    cursor++;
                }

                source = css.Substring(sourceStart, cursor - sourceStart);
                end = cursor;
            }

            int importEnd = end;
            while (importEnd < css.Length && css[importEnd] != ';') {
                importEnd++;
            }

            if (importEnd < css.Length) {
                importEnd++;
            }

            string conditionText = css.Substring(end, Math.Max(0, importEnd - end)).Trim().TrimEnd(';').Trim();
            yield return new CssImportReference(importStart, importEnd, source, conditionText);
            index = importEnd;
        }
    }

    private static bool IsApplicableCssImport(string conditionText, HtmlCssMediaContext mediaContext) {
        string remaining = conditionText.Trim();
        if (remaining.Length == 0) {
            return true;
        }

        while (remaining.Length > 0) {
            if (TryConsumeCssImportFunctionCondition(remaining, "layer", out _, out string afterLayer)) {
                remaining = afterLayer.TrimStart();
                continue;
            }

            if (StartsWithCssIdentifier(remaining, "layer")) {
                remaining = remaining.Substring("layer".Length).TrimStart();
                continue;
            }

            if (TryConsumeCssImportFunctionCondition(remaining, "supports", out string supportsCondition, out string afterSupports)) {
                if (!HtmlComputedStyleEngine.IsApplicableSupports(supportsCondition)) {
                    return false;
                }

                remaining = afterSupports.TrimStart();
                continue;
            }

            break;
        }

        return remaining.Length == 0 || IsApplicableMedia(remaining, mediaContext);
    }

    private static bool TryConsumeCssImportFunctionCondition(string text, string functionName, out string argument, out string remaining) {
        argument = string.Empty;
        remaining = text;
        if (!IsCssFunctionNameAt(text, 0, functionName)) {
            return false;
        }

        int open = text.IndexOf('(');
        if (open < 0) {
            return false;
        }

        int close = FindMatchingCssParenthesis(text, open);
        if (close <= open) {
            return false;
        }

        argument = text.Substring(open + 1, close - open - 1).Trim();
        remaining = text.Substring(close + 1);
        return true;
    }

    private static bool StartsWithCssIdentifier(string text, string identifier) {
        if (!text.StartsWith(identifier, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return text.Length == identifier.Length || !IsCssIdentifierCharacter(text[identifier.Length]);
    }

    private static bool TryReadCssUrlFunctionSource(string css, int cursor, out string source, out int end) {
        if (cursor < css.Length && (css[cursor] == '"' || css[cursor] == '\'')) {
            if (!TryReadCssQuotedValue(css, cursor, out source, out cursor)) {
                end = cursor;
                return false;
            }
        } else {
            int sourceStart = cursor;
            while (cursor < css.Length && css[cursor] != ')') {
                cursor++;
            }

            source = css.Substring(sourceStart, cursor - sourceStart).Trim();
        }

        cursor = SkipWhitespace(css, cursor);
        if (cursor < css.Length && css[cursor] == ')') {
            cursor++;
        }

        end = cursor;
        return true;
    }

    private static bool TryReadCssQuotedValue(string css, int cursor, out string value, out int end) {
        char quote = css[cursor];
        int start = cursor + 1;
        cursor = start;
        while (cursor < css.Length) {
            if (css[cursor] == quote && !IsEscaped(css, cursor)) {
                value = css.Substring(start, cursor - start);
                end = cursor + 1;
                return true;
            }

            cursor++;
        }

        value = string.Empty;
        end = cursor;
        return false;
    }

    private static int SkipWhitespace(string text, int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }

        return index;
    }

    private static bool StartsWith(string text, int index, string value) {
        return index >= 0
            && index + value.Length <= text.Length
            && string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
    }

    private static bool IsCssFunctionNameAt(string css, int index, string functionName) {
        int open = css.IndexOf('(', index);
        if (open <= index) {
            return false;
        }

        string rawName = css.Substring(index, open - index).Trim();
        if (!CssFunctionNameEquals(rawName, functionName)) {
            return false;
        }

        return index == 0 || !IsCssIdentifierCharacter(css[index - 1]);
    }

    private static bool TryFindNextCssFunction(string css, int startIndex, out int functionStart, out int open, params string[] functionNames) {
        for (open = css.IndexOf('(', Math.Max(0, startIndex)); open >= 0; open = css.IndexOf('(', open + 1)) {
            int nameEnd = open;
            int cursor = nameEnd - 1;
            while (cursor >= 0 && char.IsWhiteSpace(css[cursor])) {
                cursor--;
            }

            int trimmedEnd = cursor + 1;
            while (cursor >= 0 && (IsCssIdentifierCharacter(css[cursor]) || css[cursor] == '\\')) {
                cursor--;
            }

            int nameStart = cursor + 1;
            if (nameStart >= trimmedEnd || (nameStart > 0 && IsCssIdentifierCharacter(css[nameStart - 1]))) {
                continue;
            }

            string rawName = css.Substring(nameStart, trimmedEnd - nameStart);
            foreach (string functionName in functionNames) {
                if (CssFunctionNameEquals(rawName, functionName)) {
                    functionStart = nameStart;
                    return true;
                }
            }
        }

        functionStart = -1;
        open = -1;
        return false;
    }

    private static bool IsCssIdentifierCharacter(char value) {
        return char.IsLetterOrDigit(value)
            || value == '_'
            || value == '-'
            || value >= 0x80;
    }

    private static bool IsInsideCssString(string css, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
            }
        }

        return quote != '\0';
    }

    private static string StripCssCommentsOutsideStrings(string css) {
        var result = new System.Text.StringBuilder(css.Length);
        char quote = '\0';
        for (int i = 0; i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                result.Append(current);
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                result.Append(current);
                continue;
            }

            if (current == '/' && i + 1 < css.Length && css[i + 1] == '*') {
                i += 2;
                while (i + 1 < css.Length && !(css[i] == '*' && css[i + 1] == '/')) {
                    i++;
                }

                if (i + 1 < css.Length) {
                    i++;
                }

                result.Append(' ');
                continue;
            }

            result.Append(current);
        }

        return result.ToString();
    }

    private static bool IsCustomPropertyUrl(string css, int index) {
        return TryGetCustomPropertyName(css, index, out _);
    }

    private static bool TryGetCustomPropertyName(string css, int index, out string propertyName) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(css.LastIndexOf(';', Math.Max(0, index - 1)), blockStart);
        string declaration = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1)).TrimStart();
        if (!declaration.StartsWith("--", StringComparison.Ordinal)) {
            propertyName = string.Empty;
            return false;
        }

        int separator = declaration.IndexOf(':');
        if (separator <= 0) {
            propertyName = string.Empty;
            return false;
        }

        propertyName = declaration.Substring(0, separator).Trim();
        return propertyName.Length > 2;
    }

    private static bool IsImportAtRuleUrl(string css, int index) {
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBlockEnd = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousSemicolon, previousBlockEnd);
        string statement = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1));
        int importStart = statement.IndexOf("@import", StringComparison.OrdinalIgnoreCase);
        return importStart >= 0 && HasImportTokenBoundary(statement, importStart);
    }

    private static bool IsAtRulePreludeUrl(string css, int index) {
        int previousOpen = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousClose = css.LastIndexOf('}', Math.Max(0, index - 1));
        int previousSemicolon = css.LastIndexOf(';', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(previousOpen, Math.Max(previousClose, previousSemicolon));
        int segmentStart = Math.Max(0, previousBoundary + 1);
        string prefix = css.Substring(segmentStart, index - segmentStart);
        if (prefix.LastIndexOf('@') < 0) {
            return false;
        }

        int nextOpen = css.IndexOf('{', index);
        if (nextOpen < 0) {
            return false;
        }

        int nextSemicolon = css.IndexOf(';', index);
        int nextClose = css.IndexOf('}', index);
        return (nextSemicolon < 0 || nextOpen < nextSemicolon)
            && (nextClose < 0 || nextOpen < nextClose);
    }

    private static bool HasImportTokenBoundary(string css, int importStart) {
        return HasAtRuleTokenBoundary(css, importStart, "@import");
    }

    private static bool HasAtRuleTokenBoundary(string css, int atRuleStart, string atRuleName) {
        int afterImport = atRuleStart + atRuleName.Length;
        return afterImport >= css.Length || !IsCssIdentifierCharacter(css[afterImport]);
    }

    private static bool HasStyleRuleBefore(string css, int index) {
        char quote = '\0';
        for (int i = 0; i < index && i < css.Length; i++) {
            char current = css[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(css, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == '{' || current == '}') {
                return true;
            }
        }

        return false;
    }

    private static HtmlResourceKind ClassifyCssUrl(string css, int index) {
        string propertyName = GetCssDeclarationPropertyName(css, index);
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        string blockPrefix = blockStart >= 0 ? css.Substring(0, blockStart).ToLowerInvariant() : string.Empty;
        int fontFaceStart = blockPrefix.LastIndexOf("@font-face", StringComparison.Ordinal);
        int previousBlockEnd = blockPrefix.LastIndexOf('}');
        if (fontFaceStart >= 0 && fontFaceStart > previousBlockEnd) {
            return HtmlResourceKind.Font;
        }

        if (IsSupportedCssImageUrlProperty(propertyName)) {
            return HtmlResourceKind.Image;
        }

        return HtmlResourceKind.Other;
    }

    private static bool IsSupportedCssUrlDeclaration(string css, int index) {
        return ClassifyCssUrl(css, index) != HtmlResourceKind.Other;
    }

    private static string GetCssDeclarationPropertyName(string css, int index) {
        int declarationStart = GetDeclarationStart(css, index);
        int separator = css.IndexOf(':', declarationStart, Math.Max(0, index - declarationStart));
        if (separator <= declarationStart) {
            return string.Empty;
        }

        string propertyName = DecodeCssEscapes(css.Substring(declarationStart, separator - declarationStart).Trim());
        return propertyName.StartsWith("--", StringComparison.Ordinal)
            ? propertyName
            : propertyName.ToLowerInvariant();
    }

    private static bool IsSupportedCssImageUrlProperty(string propertyName) {
        switch (propertyName) {
            case "background":
            case "background-image":
            case "border-image":
            case "border-image-source":
            case "content":
            case "cursor":
            case "list-style":
            case "list-style-image":
            case "mask":
            case "mask-image":
            case "-webkit-mask":
            case "-webkit-mask-image":
            case "filter":
            case "clip-path":
                return true;
            default:
                return false;
        }
    }

    private static bool IsImportUrl(int index, IEnumerable<SourceRange> ranges) {
        return IsInRanges(index, ranges);
    }

    private static bool IsInRanges(int index, IEnumerable<SourceRange> ranges) {
        foreach (SourceRange range in ranges) {
            if (index >= range.Start && index < range.End) {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeSource(string source) {
        return source.Trim().Trim('\'', '"');
    }

    private static string DecodeCssEscapes(string source) {
        if (source.IndexOf('\\') < 0) {
            return source;
        }

        var result = new System.Text.StringBuilder(source.Length);
        for (int i = 0; i < source.Length; i++) {
            char current = source[i];
            if (current != '\\' || i + 1 >= source.Length) {
                result.Append(current);
                continue;
            }

            int cursor = i + 1;
            int hexStart = cursor;
            while (cursor < source.Length && cursor - hexStart < 6 && IsHexDigit(source[cursor])) {
                cursor++;
            }

            if (cursor > hexStart) {
                string hex = source.Substring(hexStart, cursor - hexStart);
                if (!int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out int codePoint)
                    || codePoint == 0
                    || codePoint > 0x10FFFF
                    || (codePoint >= 0xD800 && codePoint <= 0xDFFF)) {
                    result.Append('\uFFFD');
                } else {
                    result.Append(char.ConvertFromUtf32(codePoint));
                }

                if (cursor < source.Length && char.IsWhiteSpace(source[cursor])) {
                    cursor++;
                }

                i = cursor - 1;
                continue;
            }

            result.Append(source[cursor]);
            i = cursor;
        }

        return result.ToString();
    }

    private static bool IsExecutableScriptElement(IElement element) {
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "module", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/jscript", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsHexDigit(char value) {
        return (value >= '0' && value <= '9')
            || (value >= 'a' && value <= 'f')
            || (value >= 'A' && value <= 'F');
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static void AddSrcSet(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(raw, options.MaxResponsiveImageCandidates)) {
            AddRaw(manifest, kind, element, attributeName, candidate.Url, baseUri, options);
        }
    }

    private static void AddAttribute(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options, bool skipFragmentOnly = false) {
        string? source = element.GetAttribute(attributeName);
        if (skipFragmentOnly && IsFragmentOnlyReference(source)) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(source)) {
            AddRaw(manifest, kind, element, attributeName, source!, baseUri, options);
        }
    }

    private static bool IsFragmentOnlyReference(string? source) {
        return !string.IsNullOrWhiteSpace(source) && source!.TrimStart().StartsWith("#", StringComparison.Ordinal);
    }

    private static void AddMetaRefresh(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!string.Equals(element.GetAttribute("http-equiv"), "refresh", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        string? content = element.GetAttribute("content");
        if (string.IsNullOrWhiteSpace(content)) {
            return;
        }

        if (!TryReadMetaRefreshUrl(content!, out string source)) {
            return;
        }

        AddRaw(manifest, HtmlResourceKind.Hyperlink, element, "content", source, baseUri, options);
    }

    private static bool TryReadMetaRefreshUrl(string content, out string source) {
        source = string.Empty;
        foreach (string parameter in SplitMetaRefreshParameters(content).Skip(1)) {
            int separator = parameter.IndexOf('=');
            if (separator <= 0) {
                continue;
            }

            string name = parameter.Substring(0, separator).Trim();
            if (!string.Equals(name, "url", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            source = parameter.Substring(separator + 1).Trim();
            break;
        }

        if (source.Length == 0) {
            return false;
        }

        if (source.Length > 1 && ((source[0] == '"' && source[source.Length - 1] == '"') || (source[0] == '\'' && source[source.Length - 1] == '\''))) {
            source = source.Substring(1, source.Length - 2).Trim();
        }

        return source.Length > 0;
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

    private static void AddRaw(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, string source, Uri? baseUri, HtmlResourcePipelineOptions options) {
        HtmlUrlPolicy? policy = kind == HtmlResourceKind.Hyperlink
            ? options.UrlPolicy
            : HtmlResourceUrlPolicy.Create(options.UrlPolicy);
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
        bool isAllowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(kind, resolved);
        manifest.Add(new HtmlResourceReference(
            kind,
            element.TagName.ToLowerInvariant(),
            attributeName,
            source.Trim(),
            resolved,
            isAllowed,
            isAllowed ? string.Empty : GetDiagnosticCode(kind)));
    }

    private static string GetDiagnosticCode(HtmlResourceKind kind) {
        switch (kind) {
            case HtmlResourceKind.Image:
                return "ImageResourceRejectedByPolicy";
            case HtmlResourceKind.Stylesheet:
                return "StylesheetResourceRejectedByPolicy";
            case HtmlResourceKind.Hyperlink:
                return "HyperlinkRejectedByPolicy";
            case HtmlResourceKind.Script:
                return "ScriptResourceRejectedByPolicy";
            case HtmlResourceKind.Media:
                return "MediaResourceRejectedByPolicy";
            case HtmlResourceKind.Font:
                return "FontResourceRejectedByPolicy";
            default:
                return "HtmlResourceRejectedByPolicy";
        }
    }

    private static bool IsResourceKindSchemeAllowed(HtmlResourceKind kind, string resolved) {
        if (kind == HtmlResourceKind.Hyperlink) {
            return true;
        }

        return !Uri.TryCreate(resolved, UriKind.Absolute, out var uri)
            || !uri.Scheme.Equals(Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAllowedResourceCandidate(HtmlResourceKind kind, string? source, Uri? baseUri, HtmlUrlPolicy resourcePolicy) {
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
        return !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(kind, resolved);
    }

    private sealed class CssImportReference {
        internal CssImportReference(int start, int end, string source, string conditionText) {
            Start = start;
            End = end;
            Source = source;
            ConditionText = conditionText;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
        internal string ConditionText { get; }
    }

    private sealed class SourceRange {
        internal SourceRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }

    private sealed class CssStringUrlReference {
        internal CssStringUrlReference(int start, int end, string source) {
            Start = start;
            End = end;
            Source = source;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
    }

    private sealed class CssCustomPropertyDefinition {
        internal CssCustomPropertyDefinition(string source, string selector, int declarationStart, bool hasUrl, bool isImportant, IReadOnlyList<string> aliases, bool isInline, IElement? inlineOwner, string valueText, string? fallbackAlias) {
            Source = source;
            Selector = selector;
            DeclarationStart = declarationStart;
            HasUrl = hasUrl;
            IsImportant = isImportant;
            Aliases = aliases;
            IsInline = isInline;
            InlineOwner = inlineOwner;
            ValueText = valueText;
            FallbackAlias = fallbackAlias;
        }

        internal string Source { get; }
        internal string Selector { get; }
        internal int DeclarationStart { get; }
        internal bool HasUrl { get; }
        internal bool IsImportant { get; }
        internal IReadOnlyList<string> Aliases { get; }
        internal bool IsInline { get; }
        internal IElement? InlineOwner { get; }
        internal string ValueText { get; }
        internal string? FallbackAlias { get; }
        internal bool IsInheritedKeyword => string.Equals(ValueText, "inherit", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "unset", StringComparison.OrdinalIgnoreCase);
        internal bool IsCssWideInvalidatingKeyword => string.Equals(ValueText, "initial", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "revert", StringComparison.OrdinalIgnoreCase)
            || string.Equals(ValueText, "revert-layer", StringComparison.OrdinalIgnoreCase);
    }
}
