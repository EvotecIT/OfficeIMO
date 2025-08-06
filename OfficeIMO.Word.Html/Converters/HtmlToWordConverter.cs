using AngleSharp;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using OfficeIMO.Word;
using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html.Helpers;

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
    internal class HtmlToWordConverter {
        public async Task<WordDocument> ConvertAsync(string html, HtmlToWordOptions options) {
            if (html == null) throw new ArgumentNullException(nameof(html));
            options ??= new HtmlToWordOptions();

            var config = Configuration.Default.WithDefaultLoader();
            var context = BrowsingContext.New(config);
            var parser = context.GetService<IHtmlParser>();
            var document = await parser.ParseDocumentAsync(html);

            var wordDoc = WordDocument.Create();

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

        private struct TextFormatting {
            public bool Bold;
            public bool Italic;
            public bool Underline;

            public TextFormatting(bool bold = false, bool italic = false, bool underline = false) {
                Bold = bold;
                Italic = italic;
                Underline = underline;
            }
        }

        private static void ApplyParagraphStyleFromCss(WordParagraph paragraph, IElement element) {
            var style = CssStyleMapper.MapParagraphStyle(element.GetAttribute("style"));
            if (style.HasValue) {
                paragraph.Style = style.Value;
            }
        }

        private void ProcessNode(INode node, WordDocument doc, WordSection section, HtmlToWordOptions options,
            WordParagraph? currentParagraph, Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell) {
            if (node is IElement element) {
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
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                        }
                        break;
                    }
                    case "p": {
                        var paragraph = cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                        ApplyParagraphStyleFromCss(paragraph, element);
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                        }
                        break;
                    }
                    case "br": {
                        currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                        currentParagraph.AddBreak();
                        break;
                    }
                    case "strong":
                    case "b": {
                        var fmt = new TextFormatting(true, formatting.Italic, formatting.Underline);
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                        }
                        break;
                    }
                    case "em":
                    case "i": {
                        var fmt = new TextFormatting(formatting.Bold, true, formatting.Underline);
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                        }
                        break;
                    }
                    case "u": {
                        var fmt = new TextFormatting(formatting.Bold, formatting.Italic, true);
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell);
                        }
                        break;
                    }
                    case "a": {
                        var href = element.GetAttribute("href");
                        if (!string.IsNullOrEmpty(href)) {
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                            var uri = new Uri(href, UriKind.RelativeOrAbsolute);
                            var linkPara = currentParagraph.AddHyperLink(element.TextContent, uri);
                            if (!string.IsNullOrEmpty(options.FontFamily)) {
                                linkPara.SetFontFamily(options.FontFamily);
                            }
                        }
                        break;
                    }
                    case "ul":
                    case "ol": {
                        WordList list;
                        if (element.TagName.Equals("ul", StringComparison.OrdinalIgnoreCase)) {
                            list = cell != null ? cell.AddList(WordListStyle.Bulleted) : doc.AddListBulleted();
                        } else {
                            list = cell != null ? cell.AddList(WordListStyle.Headings111) : doc.AddListNumbered();
                        }
                        listStack.Push(list);
                        foreach (var li in element.Children.OfType<IHtmlListItemElement>()) {
                            ProcessNode(li, doc, section, options, null, listStack, formatting, cell);
                        }
                        listStack.Pop();
                        break;
                    }
                    case "li": {
                        var list = listStack.Peek();
                        int level = listStack.Count - 1;
                        var paragraph = list.AddItem("", level);
                        foreach (var child in element.ChildNodes) {
                            ProcessNode(child, doc, section, options, paragraph, listStack, formatting, cell);
                        }
                        break;
                    }
                    case "table": {
                        ProcessTable((IHtmlTableElement)element, doc, section, options, listStack, cell, currentParagraph);
                        break;
                    }
                    case "img": {
                        ProcessImage((IHtmlImageElement)element, doc);
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
                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : section.AddParagraph("");
                AddTextRun(currentParagraph, text, formatting, options);
            }
        }

        private static readonly System.Text.RegularExpressions.Regex _urlRegex = new(@"((?:https?|ftp)://[^\s]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        private static void AddTextRun(WordParagraph paragraph, string text, TextFormatting formatting, HtmlToWordOptions options) {
            int lastIndex = 0;
            foreach (System.Text.RegularExpressions.Match match in _urlRegex.Matches(text)) {
                if (match.Index > lastIndex) {
                    var segment = text.Substring(lastIndex, match.Index - lastIndex);
                    var run = paragraph.AddFormattedText(segment, formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                    if (!string.IsNullOrEmpty(options.FontFamily)) {
                        run.SetFontFamily(options.FontFamily);
                    }
                }
                var linkRun = paragraph.AddHyperLink(match.Value, new Uri(match.Value));
                ApplyFormatting(linkRun, formatting, options);
                lastIndex = match.Index + match.Length;
            }
            if (lastIndex < text.Length) {
                var run = paragraph.AddFormattedText(text.Substring(lastIndex), formatting.Bold, formatting.Italic, formatting.Underline ? UnderlineValues.Single : null);
                if (!string.IsNullOrEmpty(options.FontFamily)) {
                    run.SetFontFamily(options.FontFamily);
                }
            }
        }

        private static void ApplyFormatting(WordParagraph run, TextFormatting formatting, HtmlToWordOptions options) {
            if (formatting.Bold) run.SetBold();
            if (formatting.Italic) run.SetItalic();
            if (formatting.Underline) run.SetUnderline(UnderlineValues.Single);
            if (!string.IsNullOrEmpty(options.FontFamily)) run.SetFontFamily(options.FontFamily);
        }

        private void ProcessTable(IHtmlTableElement tableElem, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, WordParagraph? currentParagraph) {
            int rows = tableElem.Rows.Length;
            int cols = 0;
            foreach (var r in tableElem.Rows) {
                cols = Math.Max(cols, r.Cells.Length);
            }
            WordTable wordTable;
            if (cell != null) {
                wordTable = cell.AddTable(rows, cols);
            } else if (currentParagraph != null) {
                wordTable = currentParagraph.AddTableAfter(rows, cols);
            } else {
                var placeholder = section.AddParagraph("");
                wordTable = placeholder.AddTableAfter(rows, cols);
            }
            for (int r = 0; r < rows; r++) {
                var htmlRow = tableElem.Rows[r];
                for (int c = 0; c < htmlRow.Cells.Length; c++) {
                    var htmlCell = htmlRow.Cells[c];
                    var wordCell = wordTable.Rows[r].Cells[c];
                    if (wordCell.Paragraphs.Count == 1 && string.IsNullOrEmpty(wordCell.Paragraphs[0].Text)) {
                        wordCell.Paragraphs[0].Remove();
                    }
                    foreach (var child in htmlCell.ChildNodes) {
                        ProcessNode(child, doc, section, options, null, listStack, new TextFormatting(), wordCell);
                    }
                }
            }
        }

        private void ProcessImage(IHtmlImageElement img, WordDocument doc) {
            var src = img.Source;
            if (string.IsNullOrEmpty(src)) return;

            double? width = img.DisplayWidth > 0 ? img.DisplayWidth : null;
            double? height = img.DisplayHeight > 0 ? img.DisplayHeight : null;

            if (src.StartsWith("data:image", StringComparison.OrdinalIgnoreCase)) {
                var commaIndex = src.IndexOf(',');
                if (commaIndex > 0) {
                    var meta = src.Substring(5, commaIndex - 5); // e.g., image/png;base64
                    var base64 = src.Substring(commaIndex + 1);
                    var ext = "png";
                    var parts = meta.Split(new[] { ';', '/' }, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length >= 2) {
                        ext = parts[1];
                    }
                    doc.AddParagraph().AddImageFromBase64(base64, "image." + ext, width, height);
                }
            } else if (Uri.TryCreate(src, UriKind.Absolute, out var uri) && uri.IsFile) {
                doc.AddParagraph().AddImage(uri.LocalPath, width, height);
            } else if (File.Exists(src)) {
                doc.AddParagraph().AddImage(src, width, height);
            } else {
                doc.AddImageFromUrl(src, width, height);
            }
        }
    }
}