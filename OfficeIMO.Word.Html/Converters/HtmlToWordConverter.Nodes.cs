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
    internal partial class HtmlToWordConverter {
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
                            ApplyBidiIfPresent(element, paragraph);
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
                            ApplyBidiIfPresent(element, paragraph);
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
                            ApplyBidiIfPresent(element, paragraph);
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
                            ApplyBidiIfPresent(element, paragraph);
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
                                ApplyBidiIfPresent(element, para);
                                if (para == firstPara) {
                                    AddBookmarkIfPresent(element, para);
                                }
                            }
                            if (!string.IsNullOrEmpty(cite)) {
                                var noteRef = AddNoteReference(firstPara!, cite ?? string.Empty, options);
                                TryLinkNoteReference(noteRef, cite ?? string.Empty, options);
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
                                    ApplyBidiIfPresent(element, paragraph);
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
                                    ApplyBidiIfPresent(element, paragraph);
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
                                var fnRun = AddNoteReference(currentParagraph, title ?? string.Empty, options);
                                fnRun.SetCharacterStyleId("HtmlAbbr");
                                TryLinkNoteReference(fnRun, title ?? string.Empty, options);
                            }
                            break;
                        }
                    case "a": {
                            var href = element.GetAttribute("href");
                            var title = element.GetAttribute("title");
                            var target = element.GetAttribute("target");
                            var idAttr = element.GetAttribute("id");
                            var nameAttr = element.GetAttribute("name");
                            if (!string.IsNullOrEmpty(idAttr) || !string.IsNullOrEmpty(nameAttr)) {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                AddBookmarkIfPresent(element, currentParagraph);
                            }
                            if (string.IsNullOrWhiteSpace(href)) {
                                var fmt = formatting;
                                ApplySpanStyles(element, ref fmt);
                                foreach (var child in element.ChildNodes) {
                                    ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                                }
                                break;
                            }

                            var normalizedHref = NormalizeHref(href!);
                            if (IsInvalidHref(normalizedHref)) {
                                var fmt = formatting;
                                ApplySpanStyles(element, ref fmt);
                                foreach (var child in element.ChildNodes) {
                                    ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                                }
                                break;
                            }

                            if (normalizedHref.StartsWith("#", StringComparison.Ordinal)) {
                                var anchor = normalizedHref.TrimStart('#');
                                if (string.IsNullOrEmpty(anchor)) {
                                    var fmt = formatting;
                                    ApplySpanStyles(element, ref fmt);
                                    foreach (var child in element.ChildNodes) {
                                        ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                                    }
                                    break;
                                }

                                if (string.Equals(anchor, "top", StringComparison.OrdinalIgnoreCase) ||
                                    string.Equals(anchor, "_top", StringComparison.OrdinalIgnoreCase)) {
                                    anchor = "_top";
                                    if (headerFooter == null) {
                                        _pendingTopBookmark = true;
                                    }
                                }

                                if (_footnoteMap.TryGetValue(anchor, out var fnText)) {
                                    currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                    var noteRef = AddNoteReference(currentParagraph!, fnText ?? string.Empty, options);
                                    TryLinkNoteReference(noteRef, fnText ?? string.Empty, options);
                                    break;
                                }

                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                var fmtAnchor = formatting;
                                ApplySpanStyles(element, ref fmtAnchor);
                                var hasBlockAnchor = HasBlockDescendant(element);
                                WordParagraph linkParaAnchor;
                                if (!hasBlockAnchor && element.ChildNodes.Length > 0) {
                                    var tempParagraph = new WordParagraph(doc, newParagraph: true, newRun: false);
                                    _suppressAutoLinksDepth++;
                                    try {
                                        foreach (var child in element.ChildNodes) {
                                            ProcessNode(child, doc, section, options, tempParagraph, listStack, fmtAnchor, cell, headerFooter, headingList);
                                        }
                                    } finally {
                                        _suppressAutoLinksDepth--;
                                    }

                                    var runs = tempParagraph.GetRuns().ToList();
                                    linkParaAnchor = runs.Count > 0
                                        ? WordHyperLink.AddHyperLink(currentParagraph!, runs, anchor, tooltip: title ?? string.Empty)
                                        : currentParagraph!.AddHyperLink(element.TextContent, anchor);
                                } else {
                                    linkParaAnchor = currentParagraph!.AddHyperLink(element.TextContent, anchor);
                                }

                                if (!string.IsNullOrEmpty(options.FontFamily)) {
                                    linkParaAnchor.SetFontFamily(options.FontFamily!);
                                }
                                var linkAnchor = linkParaAnchor.Hyperlink;
                                if (linkAnchor != null) {
                                    if (!string.IsNullOrEmpty(title)) {
                                        linkAnchor.Tooltip = title;
                                    }
                                    if (!string.IsNullOrEmpty(target) && Enum.TryParse<TargetFrame>(target, true, out var frame)) {
                                        linkAnchor.TargetFrame = frame;
                                    }
                                }
                                break;
                            }

                            Uri? resolvedUri = null;
                            if (Uri.TryCreate(normalizedHref, UriKind.Absolute, out var absUri)) {
                                resolvedUri = absUri;
                            } else if (element.BaseUrl != null && Uri.TryCreate(element.BaseUrl.Href, UriKind.Absolute, out var baseUri)) {
                                if (Uri.TryCreate(baseUri, normalizedHref, out var relUri)) {
                                    resolvedUri = relUri;
                                }
                            } else if (Uri.TryCreate(normalizedHref, UriKind.RelativeOrAbsolute, out var relUri)) {
                                resolvedUri = relUri;
                            }

                            if (resolvedUri == null) {
                                var fmt = formatting;
                                ApplySpanStyles(element, ref fmt);
                                foreach (var child in element.ChildNodes) {
                                    ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                                }
                                break;
                            }

                            try {
                                currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                                var fmtExternal = formatting;
                                ApplySpanStyles(element, ref fmtExternal);
                                var hasBlock = HasBlockDescendant(element);
                                WordParagraph linkPara;
                                if (!hasBlock && element.ChildNodes.Length > 0) {
                                    var tempParagraph = new WordParagraph(doc, newParagraph: true, newRun: false);
                                    _suppressAutoLinksDepth++;
                                    try {
                                        foreach (var child in element.ChildNodes) {
                                            ProcessNode(child, doc, section, options, tempParagraph, listStack, fmtExternal, cell, headerFooter, headingList);
                                        }
                                    } finally {
                                        _suppressAutoLinksDepth--;
                                    }

                                    var runs = tempParagraph.GetRuns().ToList();
                                    linkPara = runs.Count > 0
                                        ? WordHyperLink.AddHyperLink(currentParagraph!, runs, resolvedUri, tooltip: title ?? string.Empty)
                                        : currentParagraph!.AddHyperLink(element.TextContent, resolvedUri);
                                } else {
                                    linkPara = currentParagraph!.AddHyperLink(element.TextContent, resolvedUri);
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
                            } catch (Exception) {
                                var fmt = formatting;
                                ApplySpanStyles(element, ref fmt);
                                foreach (var child in element.ChildNodes) {
                                    ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
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
                            } else if (url.Scheme == "file") {
                                TryLoadCssFromFileUrl(url);
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
                if (textNode.ParentElement != null) {
                    ApplyBidiIfPresent(textNode.ParentElement, currentParagraph);
                }
                AddTextRun(currentParagraph, text, formatting, options);
            }
        }
    }
}
