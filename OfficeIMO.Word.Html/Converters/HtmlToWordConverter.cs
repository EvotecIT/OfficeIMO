using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
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
        private readonly HashSet<IElement> _processedRadioInputs = new();
        private readonly List<ICssStyleRule> _cssRules = new();
        private readonly CssParser _cssParser = new();
        private readonly Dictionary<string, WordImage> _imageCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, byte[]> _remoteImageBytesCache = new(StringComparer.Ordinal);
        private readonly Dictionary<string, Exception> _remoteImageFailureCache = new(StringComparer.Ordinal);
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
        private long _remoteImageBytesFetched;
        private long _cssBytesUsed;
        private HtmlCssProcessingBudget _cssProcessingBudget = new HtmlCssProcessingBudget(null);
        private HtmlToWordOptions _options = new HtmlToWordOptions();
        private static readonly Regex _classRegex = new(@"\.([a-zA-Z0-9_-]+)", RegexOptions.Compiled);
        private static readonly HashSet<string> _blockTags = new(StringComparer.OrdinalIgnoreCase) {
            "p", "div", "section", "article", "aside", "nav", "header", "footer", "main",
            "table", "thead", "tbody", "tfoot", "tr", "td", "th",
            "ul", "ol", "li", "pre", "code", "blockquote", "figure", "figcaption",
            "h1", "h2", "h3", "h4", "h5", "h6", "address", "hr", "dd", "dt"
        };
        internal async Task<WordDocument> ConvertAsync(IHtmlDocument document, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;
            _cssProcessingBudget = new HtmlCssProcessingBudget(options.Limits);

            _context = document.Context;
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
            _processedRadioInputs.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _remoteImageBytesCache.Clear();
            _remoteImageFailureCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            _remoteImageBytesFetched = 0;
            _cssBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            await LoadConfiguredStylesheetsAsync(document, options, cancellationToken).ConfigureAwait(false);
            await LoadHeadStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await LoadBodyStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await PrefetchRemoteImagesAsync(document, options, cancellationToken).ConfigureAwait(false);

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

        internal async Task AddHtmlToBodyAsync(WordDocument doc, WordSection section, IHtmlDocument document, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;
            _cssProcessingBudget = new HtmlCssProcessingBudget(options.Limits);

            _context = document.Context;
            ValidateDocumentLimits(document, options);
            ApplyDocumentMetadata(doc, document);

            _footnoteMap.Clear();
            _endnoteMap.Clear();
            _commentMap.Clear();
            _unsupportedCssDiagnosticKeys.Clear();
            _processedRadioInputs.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _remoteImageBytesCache.Clear();
            _remoteImageFailureCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            _remoteImageBytesFetched = 0;
            _cssBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            await LoadConfiguredStylesheetsAsync(document, options, cancellationToken).ConfigureAwait(false);
            await LoadHeadStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await LoadBodyStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await PrefetchRemoteImagesAsync(document, options, cancellationToken).ConfigureAwait(false);

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

        internal async Task AddHtmlToHeaderAsync(WordDocument doc, WordHeader header, IHtmlDocument document, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            await AddHtmlToHeaderFooterAsync(doc, header, document, options, cancellationToken).ConfigureAwait(false);
        }

        internal async Task AddHtmlToFooterAsync(WordDocument doc, WordFooter footer, IHtmlDocument document, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            await AddHtmlToHeaderFooterAsync(doc, footer, document, options, cancellationToken).ConfigureAwait(false);
        }

        private async Task AddHtmlToHeaderFooterAsync(WordDocument doc, WordHeaderFooter headerFooter, IHtmlDocument document, HtmlToWordOptions options, CancellationToken cancellationToken) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;
            _options = options;
            _cssProcessingBudget = new HtmlCssProcessingBudget(options.Limits);

            _context = document.Context;
            ValidateDocumentLimits(document, options);
            ApplyDocumentMetadata(doc, document);

            _footnoteMap.Clear();
            _endnoteMap.Clear();
            _commentMap.Clear();
            _unsupportedCssDiagnosticKeys.Clear();
            _processedRadioInputs.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _remoteImageBytesCache.Clear();
            _remoteImageFailureCache.Clear();
            _cssClassStyles.Clear();
            _pendingTopBookmark = false;
            _imageBytesUsed = 0;
            _remoteImageBytesFetched = 0;
            _cssBytesUsed = 0;
            ResetAccessibilityDiagnosticsState();

            await LoadConfiguredStylesheetsAsync(document, options, cancellationToken).ConfigureAwait(false);
            await LoadHeadStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await LoadBodyStylesheetsAsync(document, cancellationToken).ConfigureAwait(false);
            await PrefetchRemoteImagesAsync(document, options, cancellationToken).ConfigureAwait(false);

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
