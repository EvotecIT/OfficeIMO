using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
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
                && IsApplicableMedia(sibling.GetAttribute("media") ?? string.Empty, options)
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
                    && IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, options)
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
                && IsApplicableMedia(sibling.GetAttribute("media") ?? string.Empty, options)
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
        HtmlUrlPolicy resourcePolicy = GetResourceUrlPolicy(options);
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
        HtmlUrlPolicy resourcePolicy = GetResourceUrlPolicy(options);
        foreach (string attribute in new[] { "src", "data-src" }) {
            if (IsAllowedResourceCandidate(HtmlResourceKind.Media, element.GetAttribute(attribute), baseUri, resourcePolicy)) {
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
        if ((isPreload || isStylesheet) && !IsApplicableMedia(element.GetAttribute("media") ?? string.Empty, options)) {
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

    private static bool IsApplicableMedia(string mediaText, HtmlResourcePipelineOptions options) {
        if (options.MediaWidth.HasValue && options.MediaHeight.HasValue) {
            return HtmlComputedStyleEngine.IsApplicableMedia(mediaText, options.MediaContext, options.MediaWidth.Value, options.MediaHeight.Value);
        }

        return IsApplicableMedia(mediaText, options.MediaContext);
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

        if (srcDocDepth >= HtmlConversionInputGuard.MaxSrcDocDepth) {
            return;
        }

        IHtmlDocument nested = HtmlDocumentParser.ParseDocument(srcdoc!);
        Uri? nestedBaseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(nested, baseUri);
        foreach (IElement nestedElement in nested.QuerySelectorAll(ResourceSelector)) {
            AddElementResources(manifest, nestedElement, nestedBaseUri, options, srcDocDepth + 1);
        }

        AddCssResources(manifest, nested, nestedBaseUri, options);
    }

}
