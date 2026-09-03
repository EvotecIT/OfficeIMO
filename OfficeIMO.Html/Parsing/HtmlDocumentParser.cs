using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace OfficeIMO.Html;

/// <summary>
/// Shared document parsing and base URI helpers for OfficeIMO HTML ingestion packages.
/// </summary>
internal static class HtmlDocumentParser {
    /// <summary>
    /// Parses an HTML fragment or document into an AngleSharp document.
    /// </summary>
    public static IHtmlDocument ParseDocument(string html) => ParseDocument(html, CancellationToken.None);

    /// <summary>
    /// Parses an HTML fragment or document into an AngleSharp document with cooperative cancellation.
    /// </summary>
    public static IHtmlDocument ParseDocument(string html, CancellationToken cancellationToken) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        cancellationToken.ThrowIfCancellationRequested();
        var parser = new HtmlParser(new HtmlParserOptions {
            IsKeepingSourceReferences = true
        });
        string normalized = NormalizeSvgHrefAttributeOrder(html, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return parser.ParseDocumentAsync(normalized, cancellationToken).GetAwaiter().GetResult();
    }

    internal static string? GetExactAttributeValue(IElement element, string name) =>
        GetExactAttribute(element, name)?.Value;

    internal static IAttr? GetExactAttribute(IElement element, string name) {
        bool xlink = name.StartsWith("xlink:", StringComparison.OrdinalIgnoreCase);
        string localName = xlink ? name.Substring("xlink:".Length) : name;
        const string xlinkNamespace = "http://www.w3.org/1999/xlink";
        return element.Attributes.FirstOrDefault(attribute =>
            string.Equals(attribute.LocalName, localName, StringComparison.OrdinalIgnoreCase) &&
            (xlink
                ? string.Equals(attribute.Prefix, "xlink", StringComparison.OrdinalIgnoreCase) ||
                  string.Equals(attribute.NamespaceUri, xlinkNamespace, StringComparison.Ordinal)
                : !string.Equals(attribute.Prefix, "xlink", StringComparison.OrdinalIgnoreCase) &&
                  !string.Equals(attribute.NamespaceUri, xlinkNamespace, StringComparison.Ordinal)));
    }

    private static string NormalizeSvgHrefAttributeOrder(string html, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (html.IndexOf("xlink:href", StringComparison.OrdinalIgnoreCase) < 0) return html;
        var replacements = new List<(int Start, int Length, string Value)>();
        var openElements = new List<SourceElement>();
        int cursor = 0;
        while (cursor < html.Length - 1) {
            cancellationToken.ThrowIfCancellationRequested();
            int markup = html.IndexOf('<', cursor);
            if (markup < 0 || markup == html.Length - 1) break;
            if (markup <= html.Length - 4 && string.CompareOrdinal(html, markup, "<!--", 0, 4) == 0) {
                int commentEnd = html.IndexOf("-->", markup + 4, StringComparison.Ordinal);
                cursor = commentEnd < 0 ? html.Length : commentEnd + 3;
                continue;
            }
            int nameStart = markup + 1;
            if (html[nameStart] is '!' or '?') {
                int declarationEnd = html.IndexOf('>', nameStart + 1);
                cursor = declarationEnd < 0 ? html.Length : declarationEnd + 1;
                continue;
            }
            if (html[nameStart] == '/') {
                int closingNameStart = nameStart + 1;
                int closingNameEnd = FindTagNameEnd(html, closingNameStart);
                string closingName = html.Substring(closingNameStart, closingNameEnd - closingNameStart);
                for (int index = openElements.Count - 1; index >= 0; index--) {
                    if (!openElements[index].Name.Equals(closingName, StringComparison.OrdinalIgnoreCase)) continue;
                    openElements.RemoveRange(index, openElements.Count - index);
                    break;
                }
                int closingEnd = FindStartTagEnd(html, closingNameEnd);
                cursor = closingEnd < 0 ? html.Length : closingEnd + 1;
                continue;
            }
            if (!IsAsciiLetter(html[nameStart])) {
                cursor = nameStart;
                continue;
            }
            int nameEnd = FindTagNameEnd(html, nameStart + 1);
            string tagName = html.Substring(nameStart, nameEnd - nameStart);
            int tagEnd = FindStartTagEnd(html, nameEnd);
            if (tagEnd < 0) break;
            if (ChildNamespace(openElements) != SourceNamespace.Html && IsForeignContentHtmlBreakout(tagName)) {
                while (openElements.Count > 0 && ChildNamespace(openElements) != SourceNamespace.Html) {
                    openElements.RemoveAt(openElements.Count - 1);
                }
            }
            SourceNamespace elementNamespace = ChildNamespace(openElements, tagName);
            if (elementNamespace == SourceNamespace.Svg &&
                (tagName.Equals("image", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("feimage", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("use", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("script", StringComparison.OrdinalIgnoreCase))) {
                string attributes = html.Substring(nameEnd, tagEnd - nameEnd);
                MatchCollection matches = Regex.Matches(
                    attributes,
                    "(?:^|[\\t\\n\\f\\r ])(?<name>xlink:href|href)[\\t\\n\\f\\r ]*=[\\t\\n\\f\\r ]*(?:\"[^\"]*\"|'[^']*'|[^\\t\\n\\f\\r \"'=<>]+)",
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
                    TimeSpan.FromMilliseconds(100));
                Match? href = matches.Cast<Match>().FirstOrDefault(match =>
                    match.Groups["name"].Value.Equals("href", StringComparison.OrdinalIgnoreCase));
                Match? xlink = matches.Cast<Match>().FirstOrDefault(match =>
                    match.Groups["name"].Value.Equals("xlink:href", StringComparison.OrdinalIgnoreCase));
                if (href != null && xlink != null && href.Index < xlink.Index) {
                    int start = nameEnd + href.Index;
                    int middleStart = start + href.Length;
                    int xlinkStart = nameEnd + xlink.Index;
                    string middle = html.Substring(middleStart, xlinkStart - middleStart);
                    replacements.Add((start, xlinkStart + xlink.Length - start, xlink.Value + middle + href.Value));
                }
            }
            bool selfClosing = IsSelfClosingTag(html, nameEnd, tagEnd);
            bool childrenUseHtml = elementNamespace == SourceNamespace.Html ||
                IsHtmlIntegrationPoint(html, nameEnd, tagEnd, tagName, elementNamespace);
            if (!selfClosing && !(elementNamespace == SourceNamespace.Html && IsHtmlVoidElement(tagName))) {
                openElements.Add(new SourceElement(tagName, elementNamespace, childrenUseHtml));
            }
            cursor = tagEnd + 1;
            if (elementNamespace == SourceNamespace.Html && tagName.Equals("plaintext", StringComparison.OrdinalIgnoreCase)) break;
            if (IsRawTextOrRcDataElement(tagName) &&
                (elementNamespace == SourceNamespace.Html ||
                 tagName.Equals("script", StringComparison.OrdinalIgnoreCase) ||
                 tagName.Equals("style", StringComparison.OrdinalIgnoreCase))) {
                int rawTextEnd = HtmlRawTextScanner.FindClosingTag(html, cursor, tagName);
                if (rawTextEnd < 0) break;
                cursor = rawTextEnd;
            }
        }
        if (replacements.Count == 0) return html;
        var output = new StringBuilder(html);
        foreach ((int start, int length, string value) in replacements.OrderByDescending(item => item.Start)) {
            cancellationToken.ThrowIfCancellationRequested();
            output.Remove(start, length).Insert(start, value);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return output.ToString();
    }

    private static int FindTagNameEnd(string html, int start) {
        int index = start;
        while (index < html.Length && html[index] != '>' && html[index] != '/' && !IsAsciiWhitespace(html[index])) index++;
        return index;
    }

    private static int FindStartTagEnd(string html, int start) {
        char quote = '\0';
        for (int index = start; index < html.Length; index++) {
            char current = html[index];
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current is '\'' or '"') quote = current;
            else if (current == '>') return index;
        }
        return -1;
    }

    private static bool IsSelfClosingTag(string html, int start, int tagEnd) {
        for (int index = tagEnd - 1; index >= start; index--) {
            if (IsAsciiWhitespace(html[index])) continue;
            return html[index] == '/';
        }
        return false;
    }

    private static SourceNamespace ChildNamespace(List<SourceElement> elements, string? tagName = null) {
        if (elements.Count == 0 || elements[elements.Count - 1].ChildrenUseHtml) {
            if (tagName?.Equals("svg", StringComparison.OrdinalIgnoreCase) == true) return SourceNamespace.Svg;
            if (tagName?.Equals("math", StringComparison.OrdinalIgnoreCase) == true) return SourceNamespace.MathMl;
            return SourceNamespace.Html;
        }
        return elements[elements.Count - 1].Namespace;
    }

    private static bool IsHtmlIntegrationPoint(
        string html,
        int attributeStart,
        int tagEnd,
        string tagName,
        SourceNamespace elementNamespace) {
        if (elementNamespace == SourceNamespace.Svg) {
            return tagName.Equals("foreignObject", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("desc", StringComparison.OrdinalIgnoreCase) ||
                tagName.Equals("title", StringComparison.OrdinalIgnoreCase);
        }
        if (elementNamespace != SourceNamespace.MathMl) return false;
        if (tagName.Equals("mi", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mo", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mn", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("ms", StringComparison.OrdinalIgnoreCase) ||
            tagName.Equals("mtext", StringComparison.OrdinalIgnoreCase)) {
            return true;
        }
        if (!tagName.Equals("annotation-xml", StringComparison.OrdinalIgnoreCase)) return false;

        string attributes = html.Substring(attributeStart, tagEnd - attributeStart);
        Match encoding = Regex.Match(
            attributes,
            "(?:^|[\\t\\n\\f\\r ])encoding[\\t\\n\\f\\r ]*=[\\t\\n\\f\\r ]*(?<value>\"[^\"]*\"|'[^']*'|[^\\t\\n\\f\\r \"'=<>]+)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant,
            TimeSpan.FromMilliseconds(100));
        if (!encoding.Success) return false;
        string value = encoding.Groups["value"].Value.Trim('\'', '"');
        value = System.Net.WebUtility.HtmlDecode(value).Trim();
        return value.Equals("text/html", StringComparison.OrdinalIgnoreCase) ||
            value.Equals("application/xhtml+xml", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsForeignContentHtmlBreakout(string tagName) => tagName.ToLowerInvariant() is
        "b" or "big" or "blockquote" or "body" or "br" or "center" or "code" or "dd" or "div" or "dl" or
        "dt" or "em" or "embed" or "h1" or "h2" or "h3" or "h4" or "h5" or "h6" or "head" or "hr" or
        "i" or "img" or "li" or "listing" or "menu" or "meta" or "nobr" or "ol" or "p" or "pre" or
        "ruby" or "s" or "small" or "span" or "strong" or "strike" or "sub" or "sup" or "table" or "tt" or
        "u" or "ul" or "var";

    private static bool IsHtmlVoidElement(string tagName) => tagName.ToLowerInvariant() is
        "area" or "base" or "br" or "col" or "embed" or "hr" or "img" or "input" or "link" or "meta" or
        "source" or "track" or "wbr";

    private static bool IsRawTextOrRcDataElement(string tagName) => tagName.ToLowerInvariant() is
        "script" or "style" or "xmp" or "iframe" or "noembed" or "noframes" or "textarea" or "title";

    private enum SourceNamespace { Html, Svg, MathMl }

    private readonly struct SourceElement {
        internal SourceElement(string name, SourceNamespace @namespace, bool childrenUseHtml) {
            Name = name;
            Namespace = @namespace;
            ChildrenUseHtml = childrenUseHtml;
        }

        internal string Name { get; }
        internal SourceNamespace Namespace { get; }
        internal bool ChildrenUseHtml { get; }
    }

    private static bool IsAsciiLetter(char value) => value is >= 'A' and <= 'Z' or >= 'a' and <= 'z';
    private static bool IsAsciiWhitespace(char value) => value is '\t' or '\n' or '\f' or '\r' or ' ';

    /// <summary>
    /// Creates a deep DOM clone so a target adapter can safely apply local transformations without reparsing text.
    /// </summary>
    public static IHtmlDocument CloneDocument(IHtmlDocument document) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return document.Clone(true) as IHtmlDocument
            ?? throw new InvalidOperationException("The HTML DOM implementation did not produce a document clone.");
    }

    /// <summary>
    /// Resolves the effective base URI from a parsed document and optional caller-provided fallback.
    /// </summary>
    public static Uri? ResolveEffectiveBaseUri(IHtmlDocument document, Uri? fallbackBaseUri) {
        if (document == null) {
            return fallbackBaseUri;
        }

        var baseElement = document.QuerySelector("base[href]");
        string? rawBaseHref = baseElement?.GetAttribute("href");
        if (rawBaseHref == null) {
            return fallbackBaseUri;
        }

        string baseHref = rawBaseHref.Trim();
        if (baseHref.Length == 0) {
            return fallbackBaseUri;
        }

        if (baseHref.StartsWith("//", StringComparison.Ordinal)) {
            return ResolveProtocolRelativeBaseUri(baseHref, fallbackBaseUri);
        }

        if (fallbackBaseUri != null && Uri.TryCreate(fallbackBaseUri, baseHref, out var resolvedFromFallback)) {
            return resolvedFromFallback;
        }

        if (!Uri.TryCreate(baseHref, UriKind.Absolute, out var absoluteBaseUri)) {
            return fallbackBaseUri;
        }

        // Uri treats rooted POSIX paths such as "/assets/" as file URIs. In HTML they are
        // origin-relative references and require a caller/page URI before they can be absolute.
        return absoluteBaseUri.IsFile
               && !baseHref.StartsWith(Uri.UriSchemeFile + ":", StringComparison.OrdinalIgnoreCase)
            ? fallbackBaseUri
            : absoluteBaseUri;
    }

    private static Uri? ResolveProtocolRelativeBaseUri(string baseHref, Uri? fallbackBaseUri) {
        string scheme = fallbackBaseUri != null
                        && (fallbackBaseUri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                            || fallbackBaseUri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
            ? fallbackBaseUri.Scheme
            : Uri.UriSchemeHttps;

        return Uri.TryCreate(scheme + ":" + baseHref, UriKind.Absolute, out var resolved)
            ? resolved
            : fallbackBaseUri;
    }

    /// <summary>
    /// Returns the document node that should be used as a converter traversal root.
    /// </summary>
    public static INode GetConversionRoot(IHtmlDocument document, bool useBodyContentsOnly) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        return useBodyContentsOnly && document.Body != null
            ? document.Body
            : (INode?)document.DocumentElement ?? document;
    }
}
