using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        /// <summary>
        /// Determines whether an import would perform HTTP I/O after applying the operation's
        /// image and stylesheet policies. This keeps the synchronous guard aligned with the
        /// converter instead of relying on a manifest built with different caller options.
        /// </summary>
        internal static bool RequiresRemoteAccess(IHtmlDocument document, HtmlToWordOptions options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (options == null) throw new ArgumentNullException(nameof(options));

            if (options.ImageProcessing == ImageProcessingMode.Embed) {
                foreach (IHtmlImageElement image in document.QuerySelectorAll("img").OfType<IHtmlImageElement>()) {
                    foreach (string candidate in EnumerateWordImageSourceCandidates(image, options)) {
                        string resolved = ResolveImageSourcePath(candidate, image, options);
                        if (IsHttpUri(resolved)
                            && IsImageSourceAllowed(resolved, options, out _)) {
                            return true;
                        }
                    }
                }
            }

            if (options.AllowDocumentStylesheetLinks) {
                Uri? documentBaseUri = ResolveDocumentBaseUri(document);
                foreach (IElement link in document.QuerySelectorAll("link[rel~='stylesheet'][href]")) {
                    if (TryResolveUri(link.GetAttribute("href"), documentBaseUri, out Uri? uri)
                        && IsAllowedRemoteStylesheet(uri, options)) {
                        return true;
                    }
                }
            }

            Uri? configuredBaseUri = ResolveDocumentBaseUri(document);
            foreach (string path in options.StylesheetPaths) {
                if (TryResolveUri(path, configuredBaseUri, out Uri? uri)
                    && IsAllowedRemoteStylesheet(uri, options)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsAllowedRemoteStylesheet(Uri uri, HtmlToWordOptions options) {
            return (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase))
                && options.AllowedStylesheetUriSchemes.Contains(uri.Scheme)
                && (options.AllowedStylesheetHosts.Count == 0 || options.AllowedStylesheetHosts.Contains(uri.Host));
        }

        private static bool TryResolveUri(string? value, Uri? baseUri, out Uri uri) {
            uri = null!;
            if (string.IsNullOrWhiteSpace(value)) return false;
            if (Uri.TryCreate(value, UriKind.Absolute, out Uri? absolute)) {
                uri = absolute;
                return true;
            }
            if (baseUri != null && Uri.TryCreate(baseUri, value, out Uri? relative)) {
                uri = relative;
                return true;
            }
            return false;
        }

        private static Uri? ResolveDocumentBaseUri(IHtmlDocument document) {
            string? baseHref = document.QuerySelector("base[href]")?.GetAttribute("href");
            if (Uri.TryCreate(baseHref, UriKind.Absolute, out Uri? explicitBase)) return explicitBase;
            if (Uri.TryCreate(document.BaseUrl?.Href, UriKind.Absolute, out Uri? documentBase)
                && !documentBase.Scheme.Equals("about", StringComparison.OrdinalIgnoreCase)) {
                return documentBase;
            }
            return null;
        }

        private static bool IsHttpUri(string? value) {
            return Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)
                && (uri.Scheme.Equals(Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase)
                    || uri.Scheme.Equals(Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase));
        }
    }
}
