using AngleSharp;
using AngleSharp.Css;
using AngleSharp.Css.Dom;
using AngleSharp.Css.Parser;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using AngleSharp.Io;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Concurrent;
using System.Globalization;
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

        private static List<WordParagraph> GetParagraphsInScope(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter) =>
            cell?.Paragraphs ?? headerFooter?.Paragraphs ?? section.Paragraphs;

        private static List<WordParagraph> GetGeneratedParagraphs(WordSection section, WordTableCell? cell, WordHeaderFooter? headerFooter, int startIndex) =>
            GetParagraphsInScope(section, cell, headerFooter).Skip(startIndex).ToList();

        private static bool ShouldReuseInitialWordSection(IElement element, WordDocument doc, WordSection section) {
            if (!string.Equals(element.GetAttribute("data-word-section"), "1", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            if (!element.ClassList.Contains("word-section")) {
                return false;
            }

            if (doc.Sections.Count != 1 || !ReferenceEquals(doc.Sections[0], section)) {
                return false;
            }

            return section.Tables.Count == 0 &&
                   section.Paragraphs.All(paragraph => string.IsNullOrWhiteSpace(paragraph.Text) && !paragraph.GetRuns().Any());
        }

        private static void ApplyContainerPageBreaksFromCss(IElement element, IReadOnlyList<WordParagraph> paragraphs) {
            if (paragraphs.Count == 0) {
                return;
            }

            if (StyleRequestsPageBreakBefore(element)) {
                paragraphs[0].PageBreakBefore = true;
            }

            if (StyleRequestsPageBreakAfter(element)) {
                AddPageBreakAfter(paragraphs[paragraphs.Count - 1]);
            }
        }

        private void ProcessNode(INode node, WordDocument doc, WordSection section, HtmlToWordOptions options,
            WordParagraph? currentParagraph, Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter = null, WordList? headingList = null) {
            if (node is IElement element) {
                if (string.Equals(element.TagName, "body", StringComparison.OrdinalIgnoreCase)) {
                    ParseLeadingStylesheetChildren(element);
                }
                ApplyCssToElement(element);
                ReportAccessibilityDiagnostics(element);
                switch (element.TagName.ToLowerInvariant()) {
                    case "body": {
                            var fmt = formatting;
                            var bodyStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(bodyStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            foreach (var child in element.ChildNodes) {
                                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
                            }
                            break;
                        }
                    case "section": {
                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            if (options.SectionTagHandling == SectionTagHandling.WordSection) {
                                var newSection = ShouldReuseInitialWordSection(element, doc, section) ? section : doc.AddSection();
                                ApplyExportedSectionMetadata(element, newSection);
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
                            if (TryProcessExportedHeaderFooterRegion(element, doc, section, options, formatting, cell, headerFooter)) break;

                            var fmt = formatting;
                            var divStyle = element.GetAttribute("style");
                            if (!string.IsNullOrWhiteSpace(divStyle)) {
                                ApplySpanStyles(element, ref fmt);
                            }
                            // Track start within this section rather than whole document
                            int startIndex = section.Paragraphs.Count;
                            int scopeStartIndex = GetParagraphsInScope(section, cell, headerFooter).Count;
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
                            ApplyContainerPageBreaksFromCss(element, GetGeneratedParagraphs(section, cell, headerFooter, scopeStartIndex));
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
                            ApplyPageBreakAfterFromCss(paragraph, element);
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
                            ApplyPageBreakAfterFromCss(paragraph, element);
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
                            ApplyPageBreakAfterFromCss(paragraph, element);
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
                            ApplyPageBreakAfterFromCss(paragraph, element);
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
                    case "pre": {
                            ProcessPreformattedElement(element, doc, section, options, currentParagraph, cell, headerFooter);
                            break;
                        }
                    case "code": {
                            ProcessInlineCodeElement(element, doc, section, options, currentParagraph, listStack, formatting, cell, headerFooter, headingList);
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
                            int startIndex = GetParagraphsInScope(section, cell, headerFooter).Count;
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
                            ApplyContainerPageBreaksFromCss(element, GetGeneratedParagraphs(section, cell, headerFooter, startIndex));
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
                            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
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

                                if (TryProcessNoteAnchor(anchor, section, options, ref currentParagraph, cell, headerFooter)) {
                                    break;
                                }

                                if (TryProcessCommentAnchor(anchor, section, ref currentParagraph, cell, headerFooter)) {
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
                    case "input":
                    case "select":
                    case "textarea": {
                            ProcessFormControl(element, section, options, currentParagraph, formatting, cell, headerFooter);
                            break;
                        }
                    case "datalist": {
                            break;
                        }
                    case "script":
                    case "template": {
                            AddDiagnostic(options, "HtmlElementSkipped", "HTML element content was skipped because it is not rendered as document content.", element.TagName.ToLowerInvariant());
                            break;
                        }
                    case "style": {
                            ParseCss(element.TextContent);
                            break;
                        }
                    case "link": {
                            ProcessLinkedStylesheetElement(element);
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
                    var language = GetElementLanguage(textNode.ParentElement);
                    if (!string.IsNullOrWhiteSpace(language)) {
                        formatting.Language = language;
                    }
                }
                AddTextRun(currentParagraph, text, formatting, options);
            }
        }

        private static void ApplyExportedSectionMetadata(IElement element, WordSection section) {
            if (!string.Equals(element.GetAttribute("data-word-section"), "1", StringComparison.OrdinalIgnoreCase) &&
                !element.ClassList.Contains("word-section")) {
                return;
            }

            var pageSizeValue = element.GetAttribute("data-page-size");
            if (Enum.TryParse<WordPageSize>(pageSizeValue, ignoreCase: true, out var pageSize) && pageSize != WordPageSize.Unknown) {
                section.PageSettings.PageSize = pageSize;
            }

            var orientationValue = element.GetAttribute("data-page-orientation");
            if (TryParsePageOrientation(orientationValue, out var orientation)) {
                section.PageOrientation = orientation;
            }

            if (TryGetUInt32Attribute(element, "data-page-width-twips", out var width)) {
                section.PageSettings.Width = width;
            }
            if (TryGetUInt32Attribute(element, "data-page-height-twips", out var height)) {
                section.PageSettings.Height = height;
            }
            if (TryGetInt32Attribute(element, "data-margin-top-twips", out var top)) {
                section.Margins.Top = top;
            }
            if (TryGetUInt32Attribute(element, "data-margin-right-twips", out var right)) {
                section.Margins.Right = right;
            }
            if (TryGetInt32Attribute(element, "data-margin-bottom-twips", out var bottom)) {
                section.Margins.Bottom = bottom;
            }
            if (TryGetUInt32Attribute(element, "data-margin-left-twips", out var left)) {
                section.Margins.Left = left;
            }
        }

        private static bool TryGetUInt32Attribute(IElement element, string name, out UInt32Value value) {
            value = 0U;
            if (!uint.TryParse(element.GetAttribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)) {
                return false;
            }

            value = parsed;
            return true;
        }

        private static bool TryGetInt32Attribute(IElement element, string name, out int value) =>
            int.TryParse(element.GetAttribute(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value);

        private static bool TryParsePageOrientation(string? value, out PageOrientationValues orientation) {
            orientation = PageOrientationValues.Portrait;
            if (string.Equals(value, "Landscape", StringComparison.OrdinalIgnoreCase)) {
                orientation = PageOrientationValues.Landscape;
                return true;
            }

            return string.Equals(value, "Portrait", StringComparison.OrdinalIgnoreCase);
        }
    }
}
