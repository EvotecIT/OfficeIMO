using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Concurrent;
using System.Net;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Use OfficeIMO.Word API methods instead of direct OpenXML manipulation
    /// 2. If OfficeIMO.Word API lacks needed functionality:
    ///    a. First check if similar functionality exists in OfficeIMO.Word
    ///    b. Consider adding new methods to OfficeIMO.Word API (in the main project)
    ///    c. Only use OpenXML directly as last resort for complex scenarios
    /// 3. Reuse existing OfficeIMO.Word helper methods and converters
    /// 4. Follow existing patterns in OfficeIMO.Word for consistency
    /// </summary>
    internal partial class HtmlToWordConverter {
        private readonly Dictionary<string, string[]> _footnoteMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string[]> _endnoteMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, HtmlCommentInfo> _commentMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _unsupportedCssDiagnosticKeys = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<ICssStyleRule> _cssRules = new();
        private readonly CssParser _cssParser = new();
        private readonly Dictionary<string, WordImage> _imageCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, WordParagraphStyles> _cssClassStyles = new(StringComparer.OrdinalIgnoreCase);
        private static readonly ConcurrentDictionary<string, ICssStyleRule[]> _stylesheetCache = new(StringComparer.OrdinalIgnoreCase);
        private IBrowsingContext? _context;
        private int _suppressAutoLinksDepth;
        private bool _pendingTopBookmark;
        private static readonly HttpClient _sharedHttpClient = new();
        private HttpClient _httpClient = _sharedHttpClient;
        private CancellationToken _cancellationToken = CancellationToken.None;
        private TimeSpan? _resourceTimeout;
        private long _imageBytesUsed;
        private HtmlToWordOptions _options = new HtmlToWordOptions();
        private static readonly Regex _classRegex = new(@"\.([a-zA-Z0-9_-]+)", RegexOptions.Compiled);
        private static readonly HashSet<string> _blockTags = new(StringComparer.OrdinalIgnoreCase) {
            "p", "div", "section", "article", "aside", "nav", "header", "footer", "main",
            "table", "thead", "tbody", "tfoot", "tr", "td", "th",
            "ul", "ol", "li", "pre", "code", "blockquote", "figure", "figcaption",
            "h1", "h2", "h3", "h4", "h5", "h6", "address", "hr", "dd", "dt"
        };
        public WordDocument Convert(string html, HtmlToWordOptions options) {
            return ConvertAsync(html, options, CancellationToken.None).GetAwaiter().GetResult();
        }

        public async Task<WordDocument> ConvertAsync(string html, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);
            ValidateDocumentLimits(document, options);

            var wordDoc = WordDocument.Create();
            if (!string.IsNullOrEmpty(options.FontFamily)) {
                var resolved = ResolveFontFamily(options.FontFamily) ?? options.FontFamily;
                wordDoc.Settings.FontFamily = resolved;
            }
            ApplyDocumentMetadata(wordDoc, document);

            _footnoteMap.Clear();
            _endnoteMap.Clear();
            _commentMap.Clear();
            _unsupportedCssDiagnosticKeys.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            foreach (var path in options.StylesheetPaths) {
                if (string.IsNullOrEmpty(path)) {
                    continue;
                }
                if (Uri.TryCreate(path, UriKind.Absolute, out var absolute)) {
                    if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps) {
                        await LoadAndParseCssAsync(context, new Url(absolute.ToString()), cancellationToken).ConfigureAwait(false);
                    } else if (absolute.Scheme == Uri.UriSchemeFile && File.Exists(absolute.LocalPath)) {
                        ParseCss(File.ReadAllText(absolute.LocalPath), absolute.LocalPath);
                    }
                } else if (document.BaseUrl != null) {
                    var url = new Url(new Url(document.BaseUrl), path);
                    if (url.Scheme == "http" || url.Scheme == "https") {
                        await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                    } else if (url.Scheme == "file") {
                        TryLoadCssFromFileUrl(url);
                    }
                } else if (File.Exists(path)) {
                    ParseCss(File.ReadAllText(path), path);
                }
            }
            foreach (var content in options.StylesheetContents) {
                if (!string.IsNullOrEmpty(content)) {
                    ParseCss(content);
                }
            }

            if (document.Head != null) {
                Uri? baseUri = null;
                if (document.BaseUrl != null && Uri.TryCreate(document.BaseUrl.Href, UriKind.Absolute, out var du)) {
                    baseUri = du;
                }

                foreach (var node in document.Head.ChildNodes) {
                    if (node is IHtmlBaseElement baseElement) {
                        if (Uri.TryCreate(baseElement.Href, UriKind.Absolute, out var bu)) {
                            baseUri = bu;
                        }
                        continue;
                    }
                    if (node is IHtmlStyleElement styleElement) {
                        ParseCss(styleElement.TextContent);
                        continue;
                    }
                    if (node is IHtmlLinkElement linkElement) {
                        var rel = linkElement.GetAttribute("rel");
                        if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                        if (!options.AllowDocumentStylesheetLinks) {
                            AddDiagnostic(options, "HtmlStylesheetLinkSkipped", "HTML stylesheet link was skipped because document-provided stylesheet links are disabled.", "link");
                            continue;
                        }

                        var hrefAttr = linkElement.GetAttribute("href");
                        var href = linkElement.Href ?? hrefAttr;
                        if (string.IsNullOrEmpty(href)) {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                            ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                            continue;
                        }

                        var url = new Url(href);
                        if (!url.IsAbsolute && baseUri != null) {
                            url = new Url(new Url(baseUri.ToString()), href);
                        }

                        if (url.Scheme == "http" || url.Scheme == "https") {
                            await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                        } else if (url.Scheme == "file") {
                            TryLoadCssFromFileUrl(url);
                        }
                    }
                }
            }

            CaptureNoteSections(document);
            CaptureCommentSections(document);

            if (options.DefaultPageSize.HasValue) {
                wordDoc.PageSettings.PageSize = options.DefaultPageSize.Value;
            }
            if (options.DefaultOrientation.HasValue) {
                wordDoc.PageOrientation = options.DefaultOrientation.Value;
            }

            var section = wordDoc.Sections.First();
            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? wordDoc.AddList(WordListStyle.Headings111) : null;
            if (document.Body != null) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessNode(document.Body, wordDoc, section, options, null, listStack, new TextFormatting(), null, null, headingList);
            }

            cancellationToken.ThrowIfCancellationRequested();
            InsertTopBookmarkIfNeeded(wordDoc);
            return wordDoc;
        }

        internal async Task AddHtmlToBodyAsync(WordDocument doc, WordSection section, string html, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);
            ValidateDocumentLimits(document, options);
            ApplyDocumentMetadata(doc, document);

            _footnoteMap.Clear();
            _endnoteMap.Clear();
            _commentMap.Clear();
            _unsupportedCssDiagnosticKeys.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            foreach (var path in options.StylesheetPaths) {
                if (string.IsNullOrEmpty(path)) {
                    continue;
                }
                if (Uri.TryCreate(path, UriKind.Absolute, out var absolute)) {
                    if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps) {
                        await LoadAndParseCssAsync(context, new Url(absolute.ToString()), cancellationToken).ConfigureAwait(false);
                    } else if (absolute.Scheme == Uri.UriSchemeFile && File.Exists(absolute.LocalPath)) {
                        ParseCss(File.ReadAllText(absolute.LocalPath), absolute.LocalPath);
                    }
                } else if (document.BaseUrl != null) {
                    var url = new Url(new Url(document.BaseUrl), path);
                    if (url.Scheme == "http" || url.Scheme == "https") {
                        await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                    } else if (url.Scheme == "file") {
                        TryLoadCssFromFileUrl(url);
                    }
                } else if (File.Exists(path)) {
                    ParseCss(File.ReadAllText(path), path);
                }
            }
            foreach (var content in options.StylesheetContents) {
                if (!string.IsNullOrEmpty(content)) {
                    ParseCss(content);
                }
            }

            if (document.Head != null) {
                Uri? baseUri = null;
                if (document.BaseUrl != null && Uri.TryCreate(document.BaseUrl.Href, UriKind.Absolute, out var du)) {
                    baseUri = du;
                }

                foreach (var node in document.Head.ChildNodes) {
                    if (node is IHtmlBaseElement baseElement) {
                        if (Uri.TryCreate(baseElement.Href, UriKind.Absolute, out var bu)) {
                            baseUri = bu;
                        }
                        continue;
                    }
                    if (node is IHtmlStyleElement styleElement) {
                        ParseCss(styleElement.TextContent);
                        continue;
                    }
                    if (node is IHtmlLinkElement linkElement) {
                        var rel = linkElement.GetAttribute("rel");
                        if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                        if (!options.AllowDocumentStylesheetLinks) {
                            AddDiagnostic(options, "HtmlStylesheetLinkSkipped", "HTML stylesheet link was skipped because document-provided stylesheet links are disabled.", "link");
                            continue;
                        }

                        var hrefAttr = linkElement.GetAttribute("href");
                        var href = linkElement.Href ?? hrefAttr;
                        if (string.IsNullOrEmpty(href)) {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                            ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                            continue;
                        }

                        var url = new Url(href);
                        if (!url.IsAbsolute && baseUri != null) {
                            url = new Url(new Url(baseUri.ToString()), href);
                        }

                        if (url.Scheme == "http" || url.Scheme == "https") {
                            await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                        } else if (url.Scheme == "file") {
                            TryLoadCssFromFileUrl(url);
                        }
                    }
                }
            }

            CaptureNoteSections(document, cancellationToken);
            CaptureCommentSections(document, cancellationToken);

            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? doc.AddList(WordListStyle.Headings111) : null;
            if (document.Body != null) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessNode(document.Body, doc, section, options, null, listStack, new TextFormatting(), null, null, headingList);
            }
            InsertTopBookmarkIfNeeded(doc);
        }

        internal async Task AddHtmlToHeaderAsync(WordDocument doc, WordHeader header, string html, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            await AddHtmlToHeaderFooterAsync(doc, header, html, options, cancellationToken).ConfigureAwait(false);
        }

        internal async Task AddHtmlToFooterAsync(WordDocument doc, WordFooter footer, string html, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            await AddHtmlToHeaderFooterAsync(doc, footer, html, options, cancellationToken).ConfigureAwait(false);
        }

        private async Task AddHtmlToHeaderFooterAsync(WordDocument doc, WordHeaderFooter headerFooter, string html, HtmlToWordOptions options, CancellationToken cancellationToken) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);
            ValidateDocumentLimits(document, options);
            ApplyDocumentMetadata(doc, document);

            _footnoteMap.Clear();
            _endnoteMap.Clear();
            _commentMap.Clear();
            _unsupportedCssDiagnosticKeys.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            foreach (var path in options.StylesheetPaths) {
                if (string.IsNullOrEmpty(path)) {
                    continue;
                }
                if (Uri.TryCreate(path, UriKind.Absolute, out var absolute)) {
                    if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps) {
                        await LoadAndParseCssAsync(context, new Url(absolute.ToString()), cancellationToken).ConfigureAwait(false);
                    } else if (absolute.Scheme == Uri.UriSchemeFile && File.Exists(absolute.LocalPath)) {
                        ParseCss(File.ReadAllText(absolute.LocalPath), absolute.LocalPath);
                    }
                } else if (document.BaseUrl != null) {
                    var url = new Url(new Url(document.BaseUrl), path);
                    if (url.Scheme == "http" || url.Scheme == "https") {
                        await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                    } else if (url.Scheme == "file") {
                        TryLoadCssFromFileUrl(url);
                    }
                } else if (File.Exists(path)) {
                    ParseCss(File.ReadAllText(path), path);
                }
            }
            foreach (var content in options.StylesheetContents) {
                if (!string.IsNullOrEmpty(content)) {
                    ParseCss(content);
                }
            }

            if (document.Head != null) {
                Uri? baseUri = null;
                if (document.BaseUrl != null && Uri.TryCreate(document.BaseUrl.Href, UriKind.Absolute, out var du)) {
                    baseUri = du;
                }

                foreach (var node in document.Head.ChildNodes) {
                    if (node is IHtmlBaseElement baseElement) {
                        if (Uri.TryCreate(baseElement.Href, UriKind.Absolute, out var bu)) {
                            baseUri = bu;
                        }
                        continue;
                    }
                    if (node is IHtmlStyleElement styleElement) {
                        ParseCss(styleElement.TextContent);
                        continue;
                    }
                    if (node is IHtmlLinkElement linkElement) {
                        var rel = linkElement.GetAttribute("rel");
                        if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                        if (!options.AllowDocumentStylesheetLinks) {
                            AddDiagnostic(options, "HtmlStylesheetLinkSkipped", "HTML stylesheet link was skipped because document-provided stylesheet links are disabled.", "link");
                            continue;
                        }

                        var hrefAttr = linkElement.GetAttribute("href");
                        var href = linkElement.Href ?? hrefAttr;
                        if (string.IsNullOrEmpty(href)) {
                            continue;
                        }

                        if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                            ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                            continue;
                        }

                        var url = new Url(href);
                        if (!url.IsAbsolute && baseUri != null) {
                            url = new Url(new Url(baseUri.ToString()), href);
                        }

                        if (url.Scheme == "http" || url.Scheme == "https") {
                            await LoadAndParseCssAsync(context, url, cancellationToken).ConfigureAwait(false);
                        } else if (url.Scheme == "file") {
                            TryLoadCssFromFileUrl(url);
                        }
                    }
                }
            }

            CaptureNoteSections(document, cancellationToken);
            CaptureCommentSections(document, cancellationToken);

            var section = doc.Sections.First();
            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? headerFooter.AddList(WordListStyle.Headings111) : null;
            if (document.Body != null) {
                cancellationToken.ThrowIfCancellationRequested();
                ProcessNode(document.Body, doc, section, options, null, listStack, new TextFormatting(), null, headerFooter, headingList);
            }
        }

    }
}
