using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>
/// Shared resource discovery and policy planning for OfficeIMO HTML workflows.
/// </summary>
public static class HtmlResourcePipeline {
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
            case "image":
                AddAttribute(manifest, HtmlResourceKind.Image, element, "href", baseUri, options);
                AddAttribute(manifest, HtmlResourceKind.Image, element, "src", baseUri, options);
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
        foreach (CssImportReference reference in ExtractCssImports(css)) {
            string source = reference.Source;
            if (!string.IsNullOrWhiteSpace(source)) {
                importRanges.Add(new SourceRange(reference.Start, reference.End));
                importedSources.Add(NormalizeSource(source));
                AddRaw(manifest, HtmlResourceKind.Stylesheet, element, attributeName + "-import", source, baseUri, options);
            }
        }

        foreach (Match match in CssUrlExpression.Matches(css)) {
            string source = match.Groups["url"].Value.Trim().Trim('\'', '"');
            if (!string.IsNullOrWhiteSpace(source)
                && !IsImportUrl(match.Index, importRanges)
                && !IsInsideCssString(css, match.Index)
                && !importedSources.Contains(NormalizeSource(source))) {
                AddRaw(manifest, ClassifyCssUrl(css, match.Index), element, attributeName + "-url", source, baseUri, options);
            }
        }
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

            int cursor = SkipWhitespace(css, importStart + 7);
            string source;
            int end;
            if (StartsWith(css, cursor, "url(")) {
                cursor = SkipWhitespace(css, cursor + 4);
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

            yield return new CssImportReference(importStart, importEnd, source);
            index = importEnd;
        }
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

    private static void AddAttribute(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? source = element.GetAttribute(attributeName);
        if (!string.IsNullOrWhiteSpace(source)) {
            AddRaw(manifest, kind, element, attributeName, source!, baseUri, options);
        }
    }

    private static void AddRaw(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, string source, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, options.UrlPolicy);
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

    private sealed class CssImportReference {
        internal CssImportReference(int start, int end, string source) {
            Start = start;
            End = end;
            Source = source;
        }

        internal int Start { get; }
        internal int End { get; }
        internal string Source { get; }
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
