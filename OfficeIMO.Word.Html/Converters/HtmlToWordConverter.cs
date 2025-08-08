using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Net;

namespace OfficeIMO.Word.Html.Converters {
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
        private IBrowsingContext? _context;
        public async Task<WordDocument> ConvertAsync(string html, HtmlToWordOptions options) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            _context = context;
            var document = await context.OpenAsync(req => req.Content(html));

            var wordDoc = WordDocument.Create();

            _footnoteMap.Clear();
            _cssRules.Clear();

            foreach (var path in options.StylesheetPaths) {
                if (string.IsNullOrEmpty(path)) {
                    continue;
                }
                if (Uri.TryCreate(path, UriKind.Absolute, out var absolute)) {
                    if (absolute.Scheme == Uri.UriSchemeHttp || absolute.Scheme == Uri.UriSchemeHttps) {
                        await LoadAndParseCssAsync(context, new Url(absolute.ToString()));
                    } else if (absolute.Scheme == Uri.UriSchemeFile && File.Exists(absolute.LocalPath)) {
                        ParseCss(File.ReadAllText(absolute.LocalPath));
                    }
                } else if (document.BaseUrl != null) {
                    var url = new Url(new Url(document.BaseUrl), path);
                    if (url.Scheme == "http" || url.Scheme == "https") {
                        await LoadAndParseCssAsync(context, url);
                    } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                        ParseCss(File.ReadAllText(url.Path));
                    }
                } else if (File.Exists(path)) {
                    ParseCss(File.ReadAllText(path));
                }
            }
            foreach (var content in options.StylesheetContents) {
                if (!string.IsNullOrEmpty(content)) {
                    ParseCss(content);
                }
            }

            if (document.Head != null) {
                foreach (var style in document.Head.QuerySelectorAll("style")) {
                    ParseCss(style.TextContent);
                }
                var baseElement = document.Head.QuerySelector("base[href]") as IHtmlBaseElement;
                Uri? baseUri = null;
                if (baseElement != null && Uri.TryCreate(baseElement.Href, UriKind.Absolute, out var bu)) {
                    baseUri = bu;
                } else if (document.BaseUrl != null && Uri.TryCreate(document.BaseUrl.Href, UriKind.Absolute, out var du)) {
                    baseUri = du;
                }
                foreach (var link in document.Head.QuerySelectorAll("link")) {
                    var rel = link.GetAttribute("rel");
                    if (!string.Equals(rel, "stylesheet", StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    var hrefAttr = link.GetAttribute("href");
                    if (string.IsNullOrEmpty(hrefAttr)) {
                        continue;
                    }

                    if (File.Exists(hrefAttr)) {
                        ParseCss(File.ReadAllText(hrefAttr));
                        continue;
                    }

                    if (baseUri != null) {
                        var combined = new Uri(baseUri, hrefAttr);
                        if (combined.Scheme == Uri.UriSchemeHttp || combined.Scheme == Uri.UriSchemeHttps) {
                            await LoadAndParseCssAsync(context, new Url(combined.ToString()));
                        } else if (combined.Scheme == Uri.UriSchemeFile && File.Exists(combined.LocalPath)) {
                            ParseCss(File.ReadAllText(combined.LocalPath));
                        }
                    }
                }
            }

            var footnoteSection = document.QuerySelector("section.footnotes");
            if (footnoteSection != null) {
                foreach (var li in footnoteSection.QuerySelectorAll("li")) {
                    var id = li.GetAttribute("id");
                    if (!string.IsNullOrEmpty(id)) {
                        _footnoteMap[id] = li.TextContent?.Trim() ?? string.Empty;
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
            foreach (var child in document.Body.ChildNodes) {
                ProcessNode(child, wordDoc, section, options, null, listStack, new TextFormatting(), null);
            }

            return wordDoc;
        }

        private async Task LoadAndParseCssAsync(IBrowsingContext context, Url url) {
            var loader = context.GetService<IResourceLoader>();
            if (loader == null) {
                return;
            }
            var request = new ResourceRequest(null, url);
            var download = loader.FetchAsync(request);
            var response = await download.Task;
            if (response.StatusCode == HttpStatusCode.OK) {
                using var reader = new StreamReader(response.Content);
                var css = await reader.ReadToEndAsync();
                ParseCss(css);
            }
        }

        private void ProcessNode(INode node, WordDocument doc, WordSection section, HtmlToWordOptions options,
            WordParagraph? currentParagraph, Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell) {
            if (node is IElement element) {
                ApplyCssToElement(element);
                switch (element.TagName.ToLowerInvariant()) {
                    case "section": {
                            var newSection = doc.AddSection();
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, newSection, options, null, listStack, formatting, null);
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
                            var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            paragraph.Style = HeadingStyleMapper.GetHeadingStyleForLevel(level);
                            ApplyParagraphStyleFromCss(paragraph, element);
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                            }
                            break;
                        }
                    case "p": {
                            var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            ApplyParagraphStyleFromCss(paragraph, element);
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                            }
                            break;
                        }
                    case "blockquote": {
                            var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            paragraph.SetStyleId("Quote");
                            paragraph.IndentationBefore = 720;
                            ApplyParagraphStyleFromCss(paragraph, element);
                            ApplyClassStyle(element, paragraph, options);
                            AddBookmarkIfPresent(element, paragraph);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                            }
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
                            for (int i = start; i < end; i++) {
                                var line = lines[i];
                                var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                                paragraph.SetStyleId("HTMLPreformatted");
                                if (!string.IsNullOrEmpty(mono)) {
                                    paragraph.SetFontFamily(mono);
                                }
                                if (!bookmarkAdded) {
                                    AddBookmarkIfPresent(element, paragraph);
                                    bookmarkAdded = true;
                                }
                                var fmt = new TextFormatting(false, false, false, null, mono);
                                AddTextRun(paragraph, line, fmt, options);
                            }
                            break;
                        }
                    case "div": {
                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            foreach (var child in element.ChildNodes) {
                                if (!string.IsNullOrWhiteSpace(divStyle) && child is IElement childElement) {
                                    var merged = MergeStyles(divStyle, childElement.GetAttribute("style"));
                                    if (!string.IsNullOrEmpty(merged)) {
                                        childElement.SetAttribute("style", merged);
                                    }
                                }
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "br": {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            currentParagraph.AddBreak();
                            break;
                        }
                    case "hr": {
                            if (cell != null) {
                                cell.AddParagraph("", true).AddHorizontalLine();
                            } else {
                                section.AddParagraph("").AddHorizontalLine();
                            }
                            break;
                        }
                    case "strong":
                    case "b": {
                            var fmt = formatting;
                            fmt.Bold = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "em":
                    case "i": {
                            var fmt = formatting;
                            fmt.Italic = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "u": {
                            var fmt = formatting;
                            fmt.Underline = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "s":
                    case "del": {
                            var fmt = formatting;
                            fmt.Strike = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "ins": {
                            var fmt = formatting;
                            fmt.Underline = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "mark": {
                            var fmt = formatting;
                            fmt.Highlight = HighlightColorValues.Yellow;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "sup": {
                            var fmt = formatting;
                            fmt.Superscript = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "sub": {
                            var fmt = formatting;
                            fmt.Subscript = true;
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "span": {
                            var fmt = formatting;
                            ApplySpanStyles(element, ref fmt);
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                            }
                            break;
                        }
                    case "a": {
                            var href = element.GetAttribute("href");
                            var title = element.GetAttribute("title");
                            var target = element.GetAttribute("target");
                            var idAttr = element.GetAttribute("id");
                            if (!string.IsNullOrEmpty(idAttr)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                                AddBookmarkIfPresent(element, currentParagraph);
                            }
                            if (!string.IsNullOrEmpty(href) && href.StartsWith("#") && _footnoteMap.TryGetValue(href.TrimStart('#'), out var fnText)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                                currentParagraph.AddFootNote(fnText);
                            } else if (!string.IsNullOrEmpty(href)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                                if (href.StartsWith("#")) {
                                    var anchor = href.TrimStart('#');
                                    var linkPara = currentParagraph.AddHyperLink(element.TextContent, anchor);
                                    if (!string.IsNullOrEmpty(options.FontFamily)) {
                                        linkPara.SetFontFamily(options.FontFamily);
                                    }
                                    var link = linkPara.Hyperlink;
                                    if (!string.IsNullOrEmpty(title)) {
                                        link.Tooltip = title;
                                    }
                                    if (!string.IsNullOrEmpty(target) && Enum.TryParse<TargetFrame>(target, true, out var frame)) {
                                        link.TargetFrame = frame;
                                    }
                                } else {
                                    var uri = new Uri(href, UriKind.RelativeOrAbsolute);
                                    var linkPara = currentParagraph.AddHyperLink(element.TextContent, uri);
                                    if (!string.IsNullOrEmpty(options.FontFamily)) {
                                        linkPara.SetFontFamily(options.FontFamily);
                                    }
                                    var link = linkPara.Hyperlink;
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
                            ProcessList(element, doc, section, options, listStack, cell, formatting);
                            break;
                        }
                    case "li": {
                            ProcessListItem((IHtmlListItemElement)element, doc, section, options, listStack, formatting, cell);
                            break;
                        }
                    case "table": {
                            ProcessTable((IHtmlTableElement)element, doc, section, options, listStack, cell, currentParagraph);
                            break;
                        }
                    case "figure": {
                            var img = element.QuerySelector("img") as IHtmlImageElement;
                            if (img != null) {
                                ProcessImage(img, doc, options);
                            }
                            var caption = element.QuerySelector("figcaption");
                            if (caption != null) {
                                ApplyCssToElement(caption);
                                var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                                paragraph.SetStyleId("Caption");
                                ApplyParagraphStyleFromCss(paragraph, caption);
                                ApplyClassStyle(caption, paragraph, options);
                                AddBookmarkIfPresent(caption, paragraph);
                                foreach (var child in caption.ChildNodes) {
                                    ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                                }
                            }
                            break;
                        }
                    case "img": {
                            ProcessImage((IHtmlImageElement)element, doc, options);
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
                                ParseCss(File.ReadAllText(hrefAttr));
                                break;
                            }

                            var url = new Url(href);
                            if (!url.IsAbsolute && element.BaseUrl != null) {
                                url = new Url(new Url(element.BaseUrl), href);
                            }

                            if (url.Scheme == "http" || url.Scheme == "https") {
                                if (_context != null) {
                                    LoadAndParseCssAsync(_context, url).GetAwaiter().GetResult();
                                }
                            } else if (url.Scheme == "file" && File.Exists(url.Path)) {
                                ParseCss(File.ReadAllText(url.Path));
                            }
                            break;
                        }
                    default: {
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, formatting, cell);
                            }
                            break;
                        }
                }
            } else if (node is IText textNode) {
                var text = textNode.Text;
                if (string.IsNullOrWhiteSpace(text)) {
                    return;
                }
                currentParagraph ??= cell != null ? cell.AddParagraph(paragraph: null, removeExistingParagraphs: true) : section.AddParagraph("");
                AddTextRun(currentParagraph, text, formatting, options);
            }
        }

        private static void AddBookmarkIfPresent(IElement element, WordParagraph paragraph) {
            var id = element.GetAttribute("id");
            if (!string.IsNullOrEmpty(id)) {
                WordBookmark.AddBookmark(paragraph, id);
            }
        }

        private static void ApplyClassStyle(IElement element, WordParagraph paragraph, HtmlToWordOptions options) {
            string? classAttr = element.GetAttribute("class");
            if (string.IsNullOrWhiteSpace(classAttr)) {
                return;
            }

            foreach (var cls in classAttr.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (options.ClassStyles.TryGetValue(cls, out var style)) {
                    paragraph.Style = style;
                    break;
                }
            }
        }

        private void ParseCss(string css) {
            if (string.IsNullOrWhiteSpace(css)) {
                return;
            }

            try {
                var sheet = _cssParser.ParseStyleSheet(css);
                foreach (var rule in sheet.Rules.OfType<ICssStyleRule>()) {
                    _cssRules.Add(rule);
                }
            } catch (Exception) {
                // ignore invalid CSS blocks
            }
        }

        private void ApplyCssToElement(IElement element) {
            if (_cssRules.Count == 0) {
                return;
            }

            var accumulated = new Dictionary<string, (string Value, Priority Specificity, bool Important)>(
                StringComparer.OrdinalIgnoreCase);
            foreach (var rule in _cssRules) {
                var selector = rule.Selector;
                if (selector != null && selector.Match(element, null)) {
                    var specificity = selector.Specificity;
                    foreach (var property in rule.Style) {
                        var name = property.Name;
                        var important = property.IsImportant;
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = (property.Value, specificity, important);
                        } else if (important) {
                            if (!existing.Important || specificity >= existing.Specificity) {
                                accumulated[name] = (property.Value, specificity, true);
                            }
                        } else if (!existing.Important && specificity >= existing.Specificity) {
                            accumulated[name] = (property.Value, specificity, false);
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
                        if (!accumulated.TryGetValue(name, out var existing)) {
                            accumulated[name] = (property.Value, Priority.Inline, important);
                        } else if (important) {
                            if (!existing.Important || Priority.Inline >= existing.Specificity) {
                                accumulated[name] = (property.Value, Priority.Inline, true);
                            }
                        } else if (!existing.Important && Priority.Inline >= existing.Specificity) {
                            accumulated[name] = (property.Value, Priority.Inline, false);
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