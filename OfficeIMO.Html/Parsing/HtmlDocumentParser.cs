using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;

namespace OfficeIMO.Html;

/// <summary>
/// Shared document parsing and base URI helpers for OfficeIMO HTML ingestion packages.
/// </summary>
public static class HtmlDocumentParser {
    /// <summary>
    /// Parses an HTML fragment or document into an AngleSharp document.
    /// </summary>
    public static IHtmlDocument ParseDocument(string html) {
        if (html == null) throw new ArgumentNullException(nameof(html));
        var parser = new HtmlParser();
        return parser.ParseDocument(html);
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

        if (fallbackBaseUri != null && Uri.TryCreate(fallbackBaseUri, baseHref, out var resolvedFromFallback)) {
            return resolvedFromFallback;
        }

        return Uri.TryCreate(baseHref, UriKind.Absolute, out var absoluteBaseUri)
            ? absoluteBaseUri
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
