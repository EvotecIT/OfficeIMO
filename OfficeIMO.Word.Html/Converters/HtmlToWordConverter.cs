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
        private readonly Dictionary<string, string> _footnoteMap = new(StringComparer.OrdinalIgnoreCase);
        private readonly List<ICssStyleRule> _cssRules = new();
        private readonly CssParser _cssParser = new();
        private readonly Dictionary<string, WordImage> _imageCache = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, WordParagraphStyles> _cssClassStyles = new(StringComparer.OrdinalIgnoreCase);
        private static readonly ConcurrentDictionary<string, ICssStyleRule[]> _stylesheetCache = new(StringComparer.OrdinalIgnoreCase);
        private IBrowsingContext? _context;
        private int _suppressAutoLinksDepth;
        private static readonly HttpClient _sharedHttpClient = new();
        private HttpClient _httpClient = _sharedHttpClient;
        private CancellationToken _cancellationToken = CancellationToken.None;
        private TimeSpan? _resourceTimeout;
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

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);

            var wordDoc = WordDocument.Create();
            if (!string.IsNullOrEmpty(options.FontFamily)) {
                var resolved = ResolveFontFamily(options.FontFamily) ?? options.FontFamily;
                wordDoc.Settings.FontFamily = resolved;
            }

            _footnoteMap.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();

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
                    } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                        ParseCss(File.ReadAllText(url.Path), url.Path);
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
                        } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                            ParseCss(File.ReadAllText(url.Path), url.Path);
                        }
                    }
                }
            }

            var footnoteSection = document.QuerySelector("section.footnotes");
            if (footnoteSection != null) {
                foreach (var li in footnoteSection.QuerySelectorAll("li")) {
                    var id = li.GetAttribute("id");
                    if (!string.IsNullOrEmpty(id)) {
                        _footnoteMap[id!] = li.TextContent?.Trim() ?? string.Empty;
                    }
                }
                footnoteSection.Remove();
            }

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
                foreach (var child in document.Body.ChildNodes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    ProcessNode(child, wordDoc, section, options, null, listStack, new TextFormatting(), null, null, headingList);
                }
            }

            cancellationToken.ThrowIfCancellationRequested();
            return wordDoc;
        }

        internal async Task AddHtmlToBodyAsync(WordDocument doc, WordSection section, string html, HtmlToWordOptions options, CancellationToken cancellationToken = default) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();
            cancellationToken.ThrowIfCancellationRequested();
            _cancellationToken = cancellationToken;
            _httpClient = options.HttpClient ?? _sharedHttpClient;
            _resourceTimeout = options.ResourceTimeout;

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);

            _footnoteMap.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();

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
                    } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                        ParseCss(File.ReadAllText(url.Path), url.Path);
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
                        } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                            ParseCss(File.ReadAllText(url.Path), url.Path);
                        }
                    }
                }
            }

            var footnoteSection = document.QuerySelector("section.footnotes");
            if (footnoteSection != null) {
                foreach (var li in footnoteSection.QuerySelectorAll("li")) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var id = li.GetAttribute("id");
                    if (!string.IsNullOrEmpty(id)) {
                        _footnoteMap[id!] = li.TextContent?.Trim() ?? string.Empty;
                    }
                }
                footnoteSection.Remove();
            }

            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? doc.AddList(WordListStyle.Headings111) : null;
            if (document.Body != null) {
                foreach (var child in document.Body.ChildNodes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    ProcessNode(child, doc, section, options, null, listStack, new TextFormatting(), null, null, headingList);
                }
            }
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

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html), cancellationToken).ConfigureAwait(false);

            _footnoteMap.Clear();
            _cssRules.Clear();
            _imageCache.Clear();
            _cssClassStyles.Clear();

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
                    } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                        ParseCss(File.ReadAllText(url.Path), url.Path);
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
                        } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                            ParseCss(File.ReadAllText(url.Path), url.Path);
                        }
                    }
                }
            }

            var footnoteSection = document.QuerySelector("section.footnotes");
            if (footnoteSection != null) {
                foreach (var li in footnoteSection.QuerySelectorAll("li")) {
                    cancellationToken.ThrowIfCancellationRequested();
                    var id = li.GetAttribute("id");
                    if (!string.IsNullOrEmpty(id)) {
                        _footnoteMap[id!] = li.TextContent?.Trim() ?? string.Empty;
                    }
                }
                footnoteSection.Remove();
            }

            var section = doc.Sections.First();
            var listStack = new Stack<WordList>();
            WordList? headingList = options.SupportsHeadingNumbering ? headerFooter.AddList(WordListStyle.Headings111) : null;
            if (document.Body != null) {
                foreach (var child in document.Body.ChildNodes) {
                    cancellationToken.ThrowIfCancellationRequested();
                    ProcessNode(child, doc, section, options, null, listStack, new TextFormatting(), null, headerFooter, headingList);
                }
            }
        }

        private async Task LoadAndParseCssAsync(IBrowsingContext context, Url url, CancellationToken cancellationToken) {
            var loader = context.GetService<IResourceLoader>();
            if (loader == null) {
                return;
            }
            var request = new ResourceRequest(null!, url);
            var download = loader.FetchAsync(request);
            var response = await download.Task.ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (response.StatusCode == HttpStatusCode.OK) {
                using var reader = new StreamReader(response.Content);
#if NET8_0_OR_GREATER
                var css = await reader.ReadToEndAsync(cancellationToken).ConfigureAwait(false);
#else
                var css = await reader.ReadToEndAsync().ConfigureAwait(false);
                cancellationToken.ThrowIfCancellationRequested();
#endif
                ParseCss(css, url.Href);
            }
        }

        private static bool HasBlockDescendant(IElement element) {
            var stack = new Stack<IElement>();
            stack.Push(element);
            while (stack.Count > 0) {
                var current = stack.Pop();
                foreach (var child in current.Children) {
                    if (_blockTags.Contains(child.TagName)) {
                        return true;
                    }
                    stack.Push(child);
                }
            }
            return false;
        }

        private void ProcessNode(INode node, WordDocument doc, WordSection section, HtmlToWordOptions options,
            WordParagraph? currentParagraph, Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter = null, WordList? headingList = null) {
            if (node is IElement element) {
                ApplyCssToElement(element);
                switch (element.TagName.ToLowerInvariant()) {
                    case "section": {
                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            if (options.SectionTagHandling == SectionTagHandling.WordSection) {
                                var newSection = doc.AddSection();
                                int startIndex = newSection.Paragraphs.Count;
                                WordParagraph? para = null;
                                foreach (var child in element.ChildNodes) {
                                    if (!string.IsNullOrWhiteSpace(divStyle) && child is IElement childElement) {
                                        var merged = MergeStyles(divStyle, childElement.GetAttribute("style"));
                                        if (!string.IsNullOrEmpty(merged)) {
                                            childElement.SetAttribute("style", merged);
                                        }
                                    }
                                    ProcessNode(child, doc, newSection, options, para, listStack, fmt, null, headerFooter, headingList);
                                    para = null;
                                }
                                var secId = element.GetAttribute("id");
                                if (!string.IsNullOrEmpty(secId)) {
                                    var paragraph = newSection.Paragraphs.Count > startIndex ? newSection.Paragraphs[startIndex] : newSection.AddParagraph("");
                                    WordBookmark.AddBookmark(paragraph, $"section:{secId}");
                                }
                            } else {
                                int startIndex = section.Paragraphs.Count;
                                WordParagraph? para = currentParagraph;
                                foreach (var child in element.ChildNodes) {
                                    if (!string.IsNullOrWhiteSpace(divStyle) && child is IElement childElement) {
                                        var merged = MergeStyles(divStyle, childElement.GetAttribute("style"));
                                        if (!string.IsNullOrEmpty(merged)) {
                                            childElement.SetAttribute("style", merged);
                                        }
                                    }
                                    ProcessNode(child, doc, section, options, para, listStack, fmt, cell, headerFooter, headingList);
                                    para = null;
                                }
                                var secId = element.GetAttribute("id");
                                if (!string.IsNullOrEmpty(secId)) {
                                    var paragraph = section.Paragraphs.Count > startIndex ? section.Paragraphs[startIndex] : (cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph(""));
                                    WordBookmark.AddBookmark(paragraph, $"section:{secId}");
                                }
                            }
                            break;
                        }
                    case "article":
                    case "aside":
                    case "nav":
                    case "header":
                    case "footer":
                    case "main": {
                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            // Track start within this section rather than whole document
                            int startIndex = section.Paragraphs.Count;
                            WordParagraph? para = currentParagraph;
                            foreach (var child in element.ChildNodes) {
                                if (!string.IsNullOrWhiteSpace(divStyle) && child is IElement childElement) {
                                    var merged = MergeStyles(divStyle, childElement.GetAttribute("style"));
                                    if (!string.IsNullOrEmpty(merged)) {
                                        childElement.SetAttribute("style", merged);
                                    }
                                }
                                ProcessNode(child, doc, section, options, para, listStack, fmt, cell, headerFooter, headingList);
                                para = null;
                            }
                            var id = element.GetAttribute("id");
                            if (!string.IsNullOrEmpty(id)) {
                                var paragraph = section.Paragraphs.Count > startIndex ? section.Paragraphs[startIndex] : (cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph(""));
                                WordBookmark.AddBookmark(paragraph, $"{element.TagName.ToLowerInvariant()}:{id}");
                            }
                            break;
                        }
                    case "h1":
                    case "h2":
                    case "h3":
                    case "h4":
                    case "h5":
                    case "h6": {
                            int level = int.Parse(element.TagName.Substring(1));
                            WordParagraph paragraph;
                            if (options.SupportsHeadingNumbering && headingList != null && cell == null) {
                                paragraph = headingList.AddItem("", level - 1);
                            } else {
                                paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            }
                            paragraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(level);
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var props = ApplyParagraphStyleFromCss(paragraph, element);
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            if (props.WhiteSpace.HasValue) {
                                fmt.WhiteSpace = props.WhiteSpace.Value;
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "p": {
                            var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var props = ApplyParagraphStyleFromCss(paragraph, element);
                            if (props.WhiteSpace.HasValue) {
                                fmt.WhiteSpace = props.WhiteSpace.Value;
                            }
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "dt": {
                            var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var props = ApplyParagraphStyleFromCss(paragraph, element);
                            if (props.WhiteSpace.HasValue) {
                                fmt.WhiteSpace = props.WhiteSpace.Value;
                            }
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "dd": {
                            var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var props = ApplyParagraphStyleFromCss(paragraph, element);
                            if (props.WhiteSpace.HasValue) {
                                fmt.WhiteSpace = props.WhiteSpace.Value;
                            }
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            var currentIndent = paragraph.IndentationBefore ?? 0;
                            if (currentIndent < 720) {
                                paragraph.IndentationBefore = 720;
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "blockquote": {
                            var startIndex = doc.Paragraphs.Count;
                            var cite = element.GetAttribute("cite");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            WordParagraph? firstPara = null;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, firstPara, listStack, fmt, cell, headerFooter, headingList);
                                if (firstPara == null && doc.Paragraphs.Count > startIndex) {
                                    firstPara = doc.Paragraphs[startIndex];
                                }
                            }
                            if (firstPara == null) {
                                firstPara = cell?.AddParagraph("", true) ?? headerFooter?.AddParagraph("") ?? section.AddParagraph("");
                            }
                            var endIndex = doc.Paragraphs.Count;
                            for (int i = startIndex; i < endIndex; i++) {
                                var para = doc.Paragraphs[i];
                                if (doc.StyleExists("Quote")) {
                                    para.SetStyleId("Quote");
                                }
                                para.IndentationBefore = 720;
                                ApplyParagraphStyleFromCss(para, element);
                                ApplyClassStyle(element, para, options);
                                if (para == firstPara) {
                                    AddBookmarkIfPresent(element, para);
                                }
                            }
                            if (!string.IsNullOrEmpty(cite)) {
                                firstPara?.AddFootNote(cite ?? string.Empty);
                            }
                            break;
                        }
                    case "svg": {
                            ProcessSvgElement(element, doc, section, options, currentParagraph, headerFooter);
                            break;
                        }
                    case "pre":
                    case "code": {
                            var textContent = element.TextContent;
                            var lines = textContent.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
                            int start = 0;
                            int end = lines.Length;
                            while (start < end && string.IsNullOrEmpty(lines[start])) start++;
                            while (end > start && string.IsNullOrEmpty(lines[end - 1])) end--;
                            var mono = FontResolver.Resolve("monospace");
                            bool bookmarkAdded = false;
                            if (options.RenderPreAsTable) {
                                WordTable preTable;
                                if (cell != null) {
                                    preTable = cell.AddTable(1, 1);
                                } else if (currentParagraph != null) {
                                    preTable = currentParagraph.AddTableAfter(1, 1);
                                } else if (headerFooter != null) {
                                    preTable = headerFooter.AddTable(1, 1);
                                } else {
                                    var placeholder = section.AddParagraph("");
                                    preTable = placeholder.AddTableAfter(1, 1);
                                }
                                var preCell = preTable.Rows[0].Cells[0];
                                for (int i = start; i < end; i++) {
                                    var line = lines[i];
                                    var paragraph = i == start ? preCell.AddParagraph("", true) : preCell.AddParagraph("");
                                    paragraph.SetStyleId("HTMLPreformatted");
                                    if (!string.IsNullOrEmpty(mono)) {
                                        paragraph.SetFontFamily(mono!);
                                    }
                                    if (!bookmarkAdded) {
                                        AddBookmarkIfPresent(element, paragraph);
                                        bookmarkAdded = true;
                                    }
                                    var fmt = new TextFormatting(false, false, false, null, mono);
                                    AddTextRun(paragraph, line, fmt, options);
                                }
                            } else {
                                for (int i = start; i < end; i++) {
                                    var line = lines[i];
                                    var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                    paragraph.SetStyleId("HTMLPreformatted");
                                    if (!string.IsNullOrEmpty(mono)) {
                                        paragraph.SetFontFamily(mono!);
                                    }
                                    if (!bookmarkAdded) {
                                        AddBookmarkIfPresent(element, paragraph);
                                        bookmarkAdded = true;
                                    }
                                    var fmt = new TextFormatting(false, false, false, null, mono);
                                    AddTextRun(paragraph, line, fmt, options);
                                }
                            }
                            break;
                        }
                    case "div":
                    case "address":
                    case "dl": {
                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            WordParagraph? para = currentParagraph;
                            foreach (var child in element.ChildNodes) {
                                if (!string.IsNullOrWhiteSpace(divStyle) && child is IElement childElement) {
                                    var merged = MergeStyles(divStyle, childElement.GetAttribute("style"));
                                    if (!string.IsNullOrEmpty(merged)) {
                                        childElement.SetAttribute("style", merged);
                                    }
                                }
                                ProcessNode(child, doc, section, options, para, listStack, fmt, cell, headerFooter, headingList);
                                if (para == null && doc.Paragraphs.Count > 0) {
                                    para = doc.Paragraphs.Last();
                                }
                            }
                            break;
                        }
                    case "br": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            currentParagraph.AddBreak();
                            break;
                        }
                    case "hr": {
                            if (cell != null) {
                                cell.AddParagraph("", true).AddHorizontalLine();
                            } else {
                                (headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("")).AddHorizontalLine();
                            }
                            break;
                        }
                    case "strong":
                    case "b": {
                            var fmt = formatting;
                            fmt.Bold = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "em":
                    case "i": {
                            var fmt = formatting;
                            fmt.Italic = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "u": {
                            var fmt = formatting;
                            fmt.Underline = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "s":
                    case "del": {
                            var fmt = formatting;
                            fmt.Strike = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "ins": {
                            var fmt = formatting;
                            fmt.Underline = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "mark": {
                            var fmt = formatting;
                            fmt.Highlight = HighlightColorValues.Yellow;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "q": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var open = currentParagraph.AddFormattedText(options.QuotePrefix, fmt.Bold, fmt.Italic, fmt.Underline ? UnderlineValues.Single : null);
                            ApplyFormatting(open, fmt, options);
                            open.SetCharacterStyleId("HtmlQuote");
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            var close = currentParagraph.AddFormattedText(options.QuoteSuffix, fmt.Bold, fmt.Italic, fmt.Underline ? UnderlineValues.Single : null);
                            ApplyFormatting(close, fmt, options);
                            close.SetCharacterStyleId("HtmlQuote");
                            break;
                        }
                    case "cite": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            fmt.Italic = true;
                            ApplySpanStyles(element, ref fmt);
                            int startRuns = currentParagraph.GetRuns().Count();
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            var runs = currentParagraph.GetRuns().ToList();
                            for (int i = startRuns; i < runs.Count; i++) {
                                runs[i].SetCharacterStyleId("HtmlCite");
                            }
                            break;
                        }
                    case "dfn": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            fmt.Italic = true;
                            ApplySpanStyles(element, ref fmt);
                            int startRuns = currentParagraph.GetRuns().Count();
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            var runs = currentParagraph.GetRuns().ToList();
                            for (int i = startRuns; i < runs.Count; i++) {
                                runs[i].SetCharacterStyleId("HtmlDfn");
                            }
                            break;
                        }
                    case "time": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            int startRuns = currentParagraph.GetRuns().Count();
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            var runs = currentParagraph.GetRuns().ToList();
                            for (int i = startRuns; i < runs.Count; i++) {
                                runs[i].SetCharacterStyleId("HtmlTime");
                            }
                            break;
                        }
                    case "sup": {
                            var fmt = formatting;
                            fmt.Superscript = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "sub": {
                            var fmt = formatting;
                            fmt.Subscript = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "font": {
                            var fmt = formatting;
                            ApplyFontStyles(element, ref fmt);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "small": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            if (!fmt.FontSize.HasValue) {
                                fmt.FontSize = 10;
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "big": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            if (!fmt.FontSize.HasValue) {
                                fmt.FontSize = 18;
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "kbd":
                    case "samp":
                    case "tt": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            var mono = FontResolver.Resolve("monospace");
                            if (!string.IsNullOrEmpty(mono)) {
                                fmt.FontFamily = mono;
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "var": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            fmt.Italic = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "nobr": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            fmt.WhiteSpace = WhiteSpaceMode.NoWrap;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "ruby":
                    case "rb":
                    case "rt":
                    case "rp": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "span": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "abbr":
                    case "acronym": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                            var title = element.GetAttribute("title");
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            if (!string.IsNullOrEmpty(title)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                var fnRun = currentParagraph.AddFootNote(title ?? string.Empty);
                                fnRun.SetCharacterStyleId("HtmlAbbr");
                            }
                            break;
                        }
                    case "a": {
                            var href = element.GetAttribute("href");
                            var title = element.GetAttribute("title");
                            var target = element.GetAttribute("target");        
                            var idAttr = element.GetAttribute("id");
                            if (!string.IsNullOrEmpty(idAttr)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                AddBookmarkIfPresent(element, currentParagraph);
                            }
                            if (href is { Length: > 0 } h1 && h1.StartsWith("#") && _footnoteMap.TryGetValue(h1.TrimStart('#'), out var fnText)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                currentParagraph!.AddFootNote(fnText ?? string.Empty);
                            } else if (!string.IsNullOrEmpty(href)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                var fmt = formatting;
                                ApplySpanStyles(element, ref fmt);
                                var h = href!;
                                var hasBlock = HasBlockDescendant(element);
                                WordParagraph linkPara;
                                if (!hasBlock && element.ChildNodes.Length > 0) {
                                    var tempParagraph = new WordParagraph(doc, newParagraph: true, newRun: false);
                                    _suppressAutoLinksDepth++;
                                    try {
                                        foreach (var child in element.ChildNodes) {
                                            ProcessNode(child, doc, section, options, tempParagraph, listStack, fmt, cell, headerFooter, headingList);
                                        }
                                    } finally {
                                        _suppressAutoLinksDepth--;
                                    }

                                    var runs = tempParagraph.GetRuns().ToList();
                                    if (runs.Count > 0) {
                                        if (h.StartsWith("#")) {
                                            var anchor = h.TrimStart('#');
                                            linkPara = WordHyperLink.AddHyperLink(currentParagraph!, runs, anchor, tooltip: title ?? string.Empty);
                                        } else {
                                            var uri = new Uri(href, UriKind.RelativeOrAbsolute);
                                            linkPara = WordHyperLink.AddHyperLink(currentParagraph!, runs, uri, tooltip: title ?? string.Empty);
                                        }
                                    } else {
                                        linkPara = h.StartsWith("#")
                                            ? currentParagraph!.AddHyperLink(element.TextContent, h.TrimStart('#'))
                                            : currentParagraph!.AddHyperLink(element.TextContent, new Uri(href, UriKind.RelativeOrAbsolute));
                                    }
                                } else {
                                    linkPara = h.StartsWith("#")
                                        ? currentParagraph!.AddHyperLink(element.TextContent, h.TrimStart('#'))
                                        : currentParagraph!.AddHyperLink(element.TextContent, new Uri(href, UriKind.RelativeOrAbsolute));
                                }

                                if (!string.IsNullOrEmpty(options.FontFamily)) {
                                    linkPara.SetFontFamily(options.FontFamily!);
                                }
                                var link = linkPara.Hyperlink;
                                if (link != null) {
                                    if (!string.IsNullOrEmpty(title)) {
                                        link.Tooltip = title;
                                    }
                                    if (!string.IsNullOrEmpty(target) && Enum.TryParse<TargetFrame>(target, true, out var frame)) {
                                        link.TargetFrame = frame;
                                    }
                                }
                            }
                            break;
                        }
                    case "ul":
                    case "ol": {
                            ProcessList(element, doc, section, options, listStack, cell, formatting, headerFooter);
                            break;
                        }
                    case "li": {
                            ProcessListItem((IHtmlListItemElement)element, doc, section, options, listStack, formatting, cell, headerFooter);
                            break;
                        }
                    case "table": {
                            ProcessTable((IHtmlTableElement)element, doc, section, options, listStack, cell, currentParagraph, headerFooter);
                            break;
                        }
                    case "figure": {
                            WordParagraph? figPara = currentParagraph;
                            foreach (var child in element.ChildNodes) {
                                if (child is IElement childEl && string.Equals(childEl.TagName, "figcaption", StringComparison.OrdinalIgnoreCase)) {
                                    ApplyCssToElement(childEl);
                                    var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                    paragraph.SetStyleId("Caption");
                                    ApplyParagraphStyleFromCss(paragraph, childEl);
                                    ApplyClassStyle(childEl, paragraph, options);
                                    AddBookmarkIfPresent(childEl, paragraph);
                                    foreach (var captionChild in childEl.ChildNodes) {
                                        ProcessNode(captionChild, doc, section, options, paragraph, listStack, formatting, cell, headerFooter, headingList);
                                    }
                                } else {
                                    ProcessNode(child, doc, section, options, figPara, listStack, formatting, cell, headerFooter, headingList);
                                    if (figPara == null && doc.Paragraphs.Count > 0) {
                                        figPara = doc.Paragraphs.Last();
                                    }
                                }
                            }
                            break;
                        }
                    case "img": {
                            ProcessImage((IHtmlImageElement)element, doc, options, currentParagraph, headerFooter);
                            break;
                        }
                    case "style": {
                            ParseCss(element.TextContent);
                            break;
                        }
                    case "link": {
                            var rel = element.GetAttribute("rel");
                            if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                                break;
                            }

                            var hrefAttr = element.GetAttribute("href");
                            var href = (element as IHtmlLinkElement)?.Href ?? hrefAttr;
                            if (string.IsNullOrEmpty(href)) {
                                break;
                            }

                            if (!string.IsNullOrEmpty(hrefAttr) && File.Exists(hrefAttr)) {
                                ParseCss(File.ReadAllText(hrefAttr), hrefAttr);
                                break;
                            }

                            var url = new Url(href);
                            if (!url.IsAbsolute && element.BaseUrl != null) {
                                url = new Url(new Url(element.BaseUrl), href);
                            }

                            if (url.Scheme == "http" || url.Scheme == "https") {
                                if (_context != null) {
                                    LoadAndParseCssAsync(_context, url, CancellationToken.None).GetAwaiter().GetResult();
                                }
                            } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                                ParseCss(File.ReadAllText(url.Path), url.Path);
                            }
                            break;
                        }
                    default: {
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, formatting, cell, headerFooter, headingList);
                            }
                            break;
                        }
                }
            } else if (node is IText textNode) {
                var text = textNode.Text;
                if (string.IsNullOrEmpty(text)) {
                    return;
                }
                if (string.IsNullOrWhiteSpace(text)) {
                    if (currentParagraph == null) {
                        return;
                    }
                    var existing = currentParagraph.Text;
                    if (!string.IsNullOrEmpty(existing)) {
                        var last = existing[existing.Length - 1];
                        if (last == ' ' || last == '\u00A0') {
                            return;
                        }
                    }
                }
                currentParagraph ??= cell != null ? cell.AddParagraph(paragraph: null, removeExistingParagraphs: true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                AddTextRun(currentParagraph, text, formatting, options);
            }
        }

        private static void AddBookmarkIfPresent(IElement element, WordParagraph paragraph) {
            var id = element.GetAttribute("id");
            if (!string.IsNullOrEmpty(id)) {
                WordBookmark.AddBookmark(paragraph, id!);
            }
        }

        private void ApplyClassStyle(IElement element, WordParagraph paragraph, HtmlToWordOptions options) {
            string? classAttr = element.GetAttribute("class");
            if (string.IsNullOrWhiteSpace(classAttr)) {
                return;
            }

            foreach (var cls in (classAttr ?? string.Empty).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (options.ClassStyles.TryGetValue(cls, out var style) || _cssClassStyles.TryGetValue(cls, out style)) {
                    paragraph.Style = style;
                    break;
                }

                var args = WordHtmlConverterExtensions.OnStyleMissing(paragraph, cls);
                if (args.Style.HasValue) {
                    paragraph.Style = args.Style.Value;
                    break;
                }
                if (!string.IsNullOrEmpty(args.StyleId)) {
                    paragraph.SetStyleId(args.StyleId!);
                    break;
                }
            }
        }

        private void ParseCss(string css, string? key = null) {
            if (string.IsNullOrWhiteSpace(css)) {
                return;
            }

            key ??= ComputeHash(css);
            if (!_stylesheetCache.TryGetValue(key, out var rules)) {
                try {
                    var sheet = _cssParser.ParseStyleSheet(css);
                    rules = sheet.Rules.OfType<ICssStyleRule>().ToArray();
                    _stylesheetCache[key] = rules;
                } catch (Exception) {
                    _stylesheetCache[key] = Array.Empty<ICssStyleRule>();
                    return;
                }
            }

            foreach (var rule in rules) {
                _cssRules.Add(rule);

                var mapped = CssStyleMapper.MapParagraphStyle(rule.Style?.CssText);
                if (mapped.HasValue) {
                    var selectorText = rule.SelectorText;
                    if (!string.IsNullOrWhiteSpace(selectorText)) {
                        foreach (var part in selectorText.Split(',')) {
                            foreach (Match match in _classRegex.Matches(part)) {
                                _cssClassStyles[match.Groups[1].Value] = mapped.Value;
                            }
                        }
                    }
                }
            }
        }

        private static string ComputeHash(string content) {
            using var sha = SHA256.Create();
            var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(content));
            return BitConverter.ToString(bytes).Replace("-", "");
        }

        private void ApplyCssToElement(IElement element) {
            if (_cssRules.Count == 0) {
                return;
            }

            var accumulated = new Dictionary<string, (string Value, Priority Specificity, bool Important, int Order)>(
                StringComparer.OrdinalIgnoreCase);
            for (int ruleIndex = 0; ruleIndex < _cssRules.Count; ruleIndex++) {
                var rule = _cssRules[ruleIndex];
                var selector = rule.Selector;
                if (selector != null && selector.Match(element, null)) {
                    var specificity = selector.Specificity;
                    foreach (var property in rule.Style) {
                        var name = property.Name;
                        var important = property.IsImportant;
                        var candidate = (Value: property.Value, Specificity: specificity, Important: important, Order: ruleIndex);
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = candidate;
                        } else if (candidate.Important != existing.Important) {
                            if (candidate.Important) {
                                accumulated[name] = candidate;
                            }
                        } else if (specificity > existing.Specificity || (specificity == existing.Specificity && ruleIndex >= existing.Order)) {
                            accumulated[name] = candidate;
                        }
                    }
                }
            }

            var inline = element.GetAttribute("style");
            if (!string.IsNullOrEmpty(inline)) {
                try {
                    var declaration = _cssParser.ParseDeclaration(inline);
                    foreach (var property in declaration) {
                        var name = property.Name;
                        var important = property.IsImportant;
                        var candidate = (Value: property.Value, Specificity: Priority.Inline, Important: important, Order: int.MaxValue);
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = candidate;
                        } else if (candidate.Important != existing.Important) {
                            if (candidate.Important) {
                                accumulated[name] = candidate;
                            }
                        } else if (Priority.Inline > existing.Specificity || (Priority.Inline == existing.Specificity && candidate.Order >= existing.Order)) {
                            accumulated[name] = candidate;
                        }
                    }
                } catch (Exception) {
                    // ignore invalid inline style
                }
            }

            if (accumulated.Count > 0) {
                var sb = new StringBuilder();
                foreach (var kvp in accumulated) {
                    sb.Append(kvp.Key).Append(':').Append(kvp.Value.Value).Append(';');
                }
                element.SetAttribute("style", sb.ToString());
            }
        }
    }
}
