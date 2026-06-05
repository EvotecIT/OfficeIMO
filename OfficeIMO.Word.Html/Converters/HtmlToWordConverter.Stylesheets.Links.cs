using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private async Task LoadConfiguredStylesheetsAsync(IDocument document, HtmlToWordOptions options, CancellationToken cancellationToken) {
            foreach (var path in options.StylesheetPaths) {
                cancellationToken.ThrowIfCancellationRequested();
                if (string.IsNullOrEmpty(path)) {
                    continue;
                }

                if (Uri.TryCreate(path, UriKind.Absolute, out var absolute)) {
                    if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps) {
                        await LoadAndParseCssAsync(new Url(absolute.ToString()), cancellationToken).ConfigureAwait(false);
                    } else if (absolute.Scheme == Uri.UriSchemeFile && File.Exists(absolute.LocalPath)) {
                        if (TryApplyStylesheetUriPolicy(absolute, absolute.ToString())) {
                            ParseCss(ReadCssFileWithLimit(absolute.LocalPath), absolute.LocalPath);
                        }
                    }
                } else if (document.BaseUrl != null) {
                    var url = new Url(new Url(document.BaseUrl), path);
                    if (url.Scheme == "http" || url.Scheme == "https") {
                        await LoadAndParseCssAsync(url, cancellationToken).ConfigureAwait(false);
                    } else if (url.Scheme == "file") {
                        TryLoadCssFromFileUrl(url);
                    }
                } else if (File.Exists(path)) {
                    if (TryApplyLocalStylesheetPolicy(path)) {
                        ParseCss(ReadCssFileWithLimit(path), path);
                    }
                }
            }

            foreach (var content in options.StylesheetContents) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!string.IsNullOrEmpty(content)) {
                    ParseCss(content);
                }
            }
        }

        private async Task LoadHeadStylesheetsAsync(IDocument document, CancellationToken cancellationToken) {
            if (document.Head == null) {
                return;
            }

            Uri? baseUri = null;
            if (document.BaseUrl != null && Uri.TryCreate(document.BaseUrl.Href, UriKind.Absolute, out var documentBaseUri)) {
                baseUri = documentBaseUri;
            }

            foreach (var node in document.Head.ChildNodes) {
                cancellationToken.ThrowIfCancellationRequested();

                if (node is IHtmlBaseElement baseElement) {
                    if (Uri.TryCreate(baseElement.Href, UriKind.Absolute, out var headBaseUri)) {
                        baseUri = headBaseUri;
                    }
                    continue;
                }

                if (node is IHtmlStyleElement styleElement) {
                    ParseCss(styleElement.TextContent);
                    continue;
                }

                if (node is IHtmlLinkElement linkElement) {
                    await ProcessStylesheetLinkElementAsync(linkElement, baseUri, cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private void ProcessLinkedStylesheetElement(IElement element) {
            ProcessStylesheetLinkElementAsync(element, baseUri: null, _cancellationToken).GetAwaiter().GetResult();
        }

        private async Task ProcessStylesheetLinkElementAsync(IElement element, Uri? baseUri, CancellationToken cancellationToken) {
            var rel = element.GetAttribute("rel");
            if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                return;
            }

            cancellationToken.ThrowIfCancellationRequested();
            var hrefAttr = element.GetAttribute("href");
            var href = (element as IHtmlLinkElement)?.Href ?? hrefAttr;
            if (!_options.AllowDocumentStylesheetLinks) {
                AddDiagnostic(_options, "HtmlStylesheetLinkSkipped", "HTML stylesheet link was skipped because document-provided stylesheet links are disabled.", string.IsNullOrEmpty(href) ? "link" : href);
                return;
            }

            if (string.IsNullOrEmpty(href)) {
                AddDiagnostic(_options, "HtmlStylesheetLinkMissingHref", "HTML stylesheet link was skipped because it does not have an href attribute.", "link");
                return;
            }

            if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                if (TryApplyLocalStylesheetPolicy(hrefAttr!)) {
                    ParseCss(ReadCssFileWithLimit(hrefAttr!), hrefAttr);
                }
                return;
            }

            var url = new Url(href);
            if (!url.IsAbsolute && baseUri != null) {
                url = new Url(new Url(baseUri.ToString()), href);
            } else if (!url.IsAbsolute && element.BaseUrl != null) {
                url = new Url(new Url(element.BaseUrl), href);
            }

            if (url.Scheme == "http" || url.Scheme == "https") {
                if (_context != null) {
                    await LoadAndParseCssAsync(url, cancellationToken).ConfigureAwait(false);
                }
            } else if (url.Scheme == "file") {
                TryLoadCssFromFileUrl(url);
            } else if (url.IsAbsolute && Uri.TryCreate(url.Href, UriKind.Absolute, out var unsupportedUri)) {
                TryApplyStylesheetUriPolicy(unsupportedUri, url.Href);
            }
        }
    }
}
