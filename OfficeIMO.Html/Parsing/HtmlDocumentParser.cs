using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace OfficeIMO.Html;

/// <summary>
/// Shared document parsing and base URI helpers for OfficeIMO HTML ingestion packages.
/// </summary>
internal static class HtmlDocumentParser {
    /// <summary>
    /// Parses an HTML fragment or document into an AngleSharp document.
    /// </summary>
    public static IHtmlDocument ParseDocument(string html) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        var parser = new HtmlParser(new HtmlParserOptions {
            IsKeepingSourceReferences = true
        });
        return parser.ParseDocument(html);
    }

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
