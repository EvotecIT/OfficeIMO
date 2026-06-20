using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Shared resource discovery and policy planning for OfficeIMO HTML workflows.
/// </summary>
public static class HtmlResourcePipeline {
    private static readonly Regex CssImportExpression = new Regex("@import\\s+(?:url\\(\\s*)?[\"']?(?<url>[^\"')\\s;]+)[\"']?\\s*\\)?", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssUrlExpression = new Regex("url\\(\\s*(?:\"(?<url>[^\"]+)\"|'(?<url>[^']+)'|(?<url>[^)]+))\\s*\\)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssCommentExpression = new Regex("/\\*.*?\\*/", RegexOptions.Singleline | RegexOptions.CultureInvariant | RegexOptions.Compiled);

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
        foreach (IElement element in document.QuerySelectorAll("[src], [srcset], [href], [data], [data-src], [data-srcset], [poster], [data-poster]")) {
            AddElementResources(manifest, element, baseUri, options);
        }

        AddCssResources(manifest, document, baseUri, options);
        return manifest;
    }

    private static void AddElementResources(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string name = element.TagName.ToLowerInvariant();
        switch (name) {
            case "img":
                AddImage(manifest, element, baseUri, options);
                break;
            case "source":
                HtmlResourceKind sourceKind = GetSourceKind(element);
                AddSrcSet(manifest, sourceKind, element, "srcset", baseUri, options);
                AddSrcSet(manifest, sourceKind, element, "data-srcset", baseUri, options);
                AddAttribute(manifest, sourceKind, element, "src", baseUri, options);
                AddAttribute(manifest, sourceKind, element, "data-src", baseUri, options);
                break;
            case "link":
                AddLink(manifest, element, baseUri, options);
                break;
            case "base":
                break;
            case "a":
            case "area":
                AddAttribute(manifest, HtmlResourceKind.Hyperlink, element, "href", baseUri, options);
                break;
            case "script":
                AddAttribute(manifest, HtmlResourceKind.Script, element, "src", baseUri, options);
                break;
            case "video":
                AddAttribute(manifest, HtmlResourceKind.Media, element, "poster", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Media, element, "data-poster", baseUri, options);
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
            default:
                AddAttribute(manifest, HtmlResourceKind.Other, element, "src", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Other, element, "href", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Other, element, "data", baseUri, options);
                break;
        }
    }

    private static void AddImage(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        foreach (string attribute in new[] { "data-src", "src" }) {
            AddAttribute(manifest, HtmlResourceKind.Image, element, attribute, baseUri, options);
        }

        AddSrcSet(manifest, HtmlResourceKind.Image, element, "srcset", baseUri, options);
        AddSrcSet(manifest, HtmlResourceKind.Image, element, "data-srcset", baseUri, options);
    }

    private static HtmlResourceKind GetSourceKind(IElement element) {
        string parentName = element.ParentElement?.TagName.ToLowerInvariant() ?? string.Empty;
        switch (parentName) {
            case "picture":
                return HtmlResourceKind.Image;
            case "audio":
            case "video":
                return HtmlResourceKind.Media;
            default:
                return HtmlResourceKind.Other;
        }
    }

    private static void AddLink(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string rel = element.GetAttribute("rel") ?? string.Empty;
        HtmlResourceKind kind;
        if (rel.IndexOf("stylesheet", StringComparison.OrdinalIgnoreCase) >= 0) {
            kind = HtmlResourceKind.Stylesheet;
        } else if (rel.IndexOf("modulepreload", StringComparison.OrdinalIgnoreCase) >= 0) {
            kind = HtmlResourceKind.Script;
        } else if (rel.IndexOf("preload", StringComparison.OrdinalIgnoreCase) >= 0) {
            kind = GetPreloadKind(element.GetAttribute("as"));
        } else if (rel.IndexOf("font", StringComparison.OrdinalIgnoreCase) >= 0) {
            kind = HtmlResourceKind.Font;
        } else if (rel.IndexOf("icon", StringComparison.OrdinalIgnoreCase) >= 0) {
            kind = HtmlResourceKind.Image;
        } else {
            kind = HtmlResourceKind.Hyperlink;
        }

        AddAttribute(manifest, kind, element, "href", baseUri, options);
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
        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            AddCssReferences(manifest, styleElement, "css", styleElement.TextContent, baseUri, options);
        }

        foreach (IElement element in document.QuerySelectorAll("[style]")) {
            AddCssReferences(manifest, element, "style", element.GetAttribute("style") ?? string.Empty, baseUri, options);
        }
    }

    private static void AddCssReferences(HtmlResourceManifest manifest, IElement element, string attributeName, string css, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (string.IsNullOrWhiteSpace(css)) {
            return;
        }

        css = CssCommentExpression.Replace(css, string.Empty);
        var importRanges = new List<SourceRange>();
        var importedSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match match in CssImportExpression.Matches(css)) {
            string source = match.Groups["url"].Value;
            if (!string.IsNullOrWhiteSpace(source)) {
                importRanges.Add(new SourceRange(match.Index, match.Index + match.Length));
                importedSources.Add(NormalizeSource(source));
                AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", source, baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            string source = match.Groups["url"].Value.Trim().Trim('\'', '"');
            if (!string.IsNullOrWhiteSpace(source) && !IsImportUrl(match.Index, importRanges) && !importedSources.Contains(NormalizeSource(source))) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", source, baseUri, options);
            }
        }
    }

    private static HtmlResourceKind ClassifyCssUrl(string css, int index) {
        int blockStart = css.LastIndexOf('{', Math.Max(0, index - 1));
        int previousBoundary = Math.Max(css.LastIndexOf(';', Math.Max(0, index - 1)), blockStart);
        string declaration = css.Substring(Math.Max(0, previousBoundary + 1), index - Math.Max(0, previousBoundary + 1)).ToLowerInvariant();
        string blockPrefix = blockStart >= 0 ? css.Substring(0, blockStart).ToLowerInvariant() : string.Empty;
        int fontFaceStart = blockPrefix.LastIndexOf("@font-face", StringComparison.Ordinal);
        int previousBlockEnd = blockPrefix.LastIndexOf('}');
        if (fontFaceStart >= 0 && fontFaceStart > previousBlockEnd) {
            return HtmlResourceKind.Font;
        }

        if (declaration.IndexOf("background", StringComparison.Ordinal) >= 0
            || declaration.IndexOf("image", StringComparison.Ordinal) >= 0
            || declaration.IndexOf("cursor", StringComparison.Ordinal) >= 0
            || declaration.IndexOf("list-style", StringComparison.Ordinal) >= 0) {
            return HtmlResourceKind.Image;
        }

        return HtmlResourceKind.Other;
    }

    private static bool IsImportUrl(int index, IEnumerable<SourceRange> ranges) {
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

    private static void AddSrcSet(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(raw, options.MaxResponsiveImageCandidates)) {
            AddRaw(manifest, kind, element, attributeName, candidate.Url, baseUri, options);
        }
    }

    private static void AddAttribute(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? source = element.GetAttribute(attributeName);
        if (!string.IsNullOrWhiteSpace(source)) {
            AddRaw(manifest, kind, element, attributeName, source!, baseUri, options);
        }
    }

    private static void AddRaw(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, string source, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, options.UrlPolicy);
        bool isAllowed = !string.IsNullOrWhiteSpace(resolved);
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

    private sealed class SourceRange {
        internal SourceRange(int start, int end) {
            Start = start;
            End = end;
        }

        internal int Start { get; }
        internal int End { get; }
    }
}
