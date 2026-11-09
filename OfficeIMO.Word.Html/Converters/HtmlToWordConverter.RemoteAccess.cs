using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Core.Internal;
using OfficeIMO.Html;
using System.Net.Http;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static readonly Lazy<HttpClient> _sharedHttpClient = new(CreateSharedHttpClient);
        private const int MaximumRemoteResourceRedirects = 50;
        private OfficeHttpCookieScope? _remoteResourceCookies;

        private static HttpClient CreateSharedHttpClient() {
            var handler = new HttpClientHandler { AllowAutoRedirect = false };
            OfficeHttpCookieScope.DisableAutomaticCookies(handler);
            return new HttpClient(handler);
        }

        private async Task<HttpResponseMessage> FetchRemoteResourceAsync(Uri uri,
            HashSet<string> allowedSchemes, HashSet<string> allowedHosts, CancellationToken cancellationToken,
            HtmlUrlPolicy? urlPolicy = null) {
            // A supplied client owns its cookies, transport and redirect behavior.
            if (_options.HttpClient is { } client) {
                using var request = new HttpRequestMessage(HttpMethod.Get, uri);
                return await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
            }

            // Preserve the shared client's timeout across the entire redirect chain.
            HttpClient sharedClient = _sharedHttpClient.Value;
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(sharedClient.Timeout);
            OfficeHttpCookieScope cookies = LazyInitializer.EnsureInitialized(
                ref _remoteResourceCookies, static () => new OfficeHttpCookieScope())!;
            for (int redirects = 0; ; redirects++) {
                using var request = new HttpRequestMessage(HttpMethod.Get, uri);
                HttpResponseMessage response = await cookies.SendAsync(
                    sharedClient, request, uri, timeout.Token).ConfigureAwait(false);
                int status = (int)response.StatusCode;
                bool redirect = status is 300 or 301 or 302 or 303 or 307;
#if !NETFRAMEWORK
                redirect |= status == 308;
#endif
                if (!redirect || redirects >= MaximumRemoteResourceRedirects) {
                    return response;
                }

                Uri? next;
                try {
                    Uri? location = response.Headers.Location;
                    if (location == null) {
                        return response;
                    }

                    next = ResolveResourceRedirect(uri, location, urlPolicy);
                    if (next == null || (next.Scheme != Uri.UriSchemeHttp && next.Scheme != Uri.UriSchemeHttps)
#if !NETFRAMEWORK
                        || (uri.Scheme == Uri.UriSchemeHttps && next.Scheme == Uri.UriSchemeHttp)
#endif
                        || !string.IsNullOrEmpty(next.UserInfo)
                        || !allowedSchemes.Contains(next.Scheme)
                        || (allowedHosts.Count > 0 && !allowedHosts.Contains(next.Host))) {
                        throw new HtmlResourcePolicyException("The redirect destination is not allowed by the resource policy.");
                    }
                } catch {
                    response.Dispose();
                    throw;
                }

                response.Dispose();
                uri = next;
            }
        }

        private static Uri? ResolveResourceRedirect(Uri current, Uri location, HtmlUrlPolicy? policy) {
            if (!Uri.TryCreate(current, location, out Uri? target)) return null;
            if (policy == null) return target;
            if (!HtmlUrlPolicyEvaluator.IsAllowed(location.OriginalString, policy)
                || !HtmlUrlPolicyEvaluator.IsAllowed(target.AbsoluteUri, policy)) return null;
            if (policy.ResolvedUrlTransform == null) return target;

            // Unlike AbsoluteUri, this representation retains a scoped IPv6 address's interface.
            string destination = target.GetComponents(UriComponents.SerializationInfoString, UriFormat.UriEscaped);
            string transformed = HtmlUrlPolicyEvaluator.ResolveUrl(destination, null, policy);
            return Uri.TryCreate(transformed, UriKind.Absolute, out Uri? rewritten) ? rewritten : null;
        }

        /// <summary>
        /// Determines whether an import would perform HTTP I/O after applying the operation's
        /// image and stylesheet policies. This keeps the synchronous guard aligned with the
        /// converter instead of relying on a manifest built with different caller options.
        /// </summary>
        internal static bool RequiresRemoteAccess(IHtmlDocument document, HtmlWordExportOptions options) {
            ArgumentNullException.ThrowIfNull(document);
            ArgumentNullException.ThrowIfNull(options);

            if (options.ImageProcessing == WordHtmlImageProcessingMode.Embed) {
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

            Uri? configuredBaseUri = ResolveDocumentBaseUri(document, options.BasePath);
            foreach (string path in options.StylesheetPaths) {
                if (TryResolveUri(path, configuredBaseUri, out Uri? uri)
                    && IsAllowedRemoteStylesheet(uri, options)) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsAllowedRemoteStylesheet(Uri uri, HtmlWordExportOptions options) {
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

        private static Uri? ResolveDocumentBaseUri(IDocument document, string? basePath = null) {
            string? baseHref = document.QuerySelector("base[href]")?.GetAttribute("href");
            if (Uri.TryCreate(baseHref, UriKind.Absolute, out Uri? explicitBase)) return explicitBase;
            if (Uri.TryCreate(document.BaseUrl?.Href, UriKind.Absolute, out Uri? documentBase)
                && !documentBase.Scheme.Equals("about", StringComparison.OrdinalIgnoreCase)) {
                return documentBase;
            }
            if (!string.IsNullOrEmpty(basePath)) {
                string directory = Path.GetFullPath(basePath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;
                return new Uri(directory, UriKind.Absolute);
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
