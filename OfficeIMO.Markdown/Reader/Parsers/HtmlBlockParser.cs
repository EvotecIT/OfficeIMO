using System.Collections.Generic;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class HtmlBlockParser : IMarkdownBlockParser {
        private enum HtmlBlockKind {
            None,
            Type1,
            Type2,
            Type3,
            Type4,
            Type5,
            Type6,
            Type7,
        }

        private readonly struct HtmlBlockState {
            public HtmlBlockState(HtmlBlockKind kind, string? primaryEndToken, string? primaryTagName, bool allowsBlankLineContinuation) {
                Kind = kind;
                PrimaryEndToken = primaryEndToken;
                PrimaryTagName = primaryTagName;
                AllowsBlankLineContinuation = allowsBlankLineContinuation;
            }

            public HtmlBlockKind Kind { get; }
            public string? PrimaryEndToken { get; }
            public string? PrimaryTagName { get; }
            public bool AllowsBlankLineContinuation { get; }

            public bool EndsOnBlankLine => Kind is HtmlBlockKind.Type6 or HtmlBlockKind.Type7;

            public bool TracksStack => !string.IsNullOrEmpty(PrimaryTagName);

            public bool IsSatisfiedBy(string line) {
                switch (Kind) {
                    case HtmlBlockKind.Type1:
                        return ContainsToken(line, PrimaryEndToken!);
                    case HtmlBlockKind.Type2:
                        return line.IndexOf("-->", StringComparison.Ordinal) >= 0;
                    case HtmlBlockKind.Type3:
                        return line.IndexOf("?>", StringComparison.Ordinal) >= 0;
                    case HtmlBlockKind.Type4:
                        return line.IndexOf('>') >= 0;
                    case HtmlBlockKind.Type5:
                        return line.IndexOf("]]>", StringComparison.Ordinal) >= 0;
                    default:
                        return false;
                }
            }

            private static bool ContainsToken(string line, string token) {
                return line.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
            }
        }

        private static readonly HashSet<string> s_BlockTags = new(StringComparer.OrdinalIgnoreCase) {
            "address", "article", "aside", "base", "basefont", "blockquote", "body", "caption", "center",
            "col", "colgroup", "dd", "details", "dialog", "dir", "div", "dl", "dt", "fieldset", "figcaption",
            "figure", "footer", "form", "frame", "frameset", "h1", "h2", "h3", "h4", "h5", "h6", "head",
            "header", "hr", "html", "iframe", "legend", "li", "link", "main", "menu", "menuitem", "nav",
            "noframes", "ol", "optgroup", "option", "p", "param", "section", "summary", "table", "tbody",
            "td", "tfoot", "th", "thead", "title", "tr", "track", "ul",
        };

        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.HtmlBlocks) return false;
            if (lines == null || lines.Length == 0) return false;

            var line = lines[i];
            if (string.IsNullOrEmpty(line)) return false;

            int indent = CountLeadingSpaces(line);
            if (indent < 0 || indent > 3) return false;

            string trimmed = indent == 0 ? line : line.Substring(indent);
            if (trimmed.Length == 0 || trimmed[0] != '<') return false;

            if (IsHeaderlessSingleRowTableMarker(trimmed)) {
                return TryParseHeaderlessSingleRowTable(lines, ref i, options, doc, state);
            }

            // Avoid treating angle-bracket autolinks like "<https://...>" as HTML blocks.
            if (TryParseAngleAutolink(trimmed, 0, out _, out _, out _)) return false;

            if (!TryGetHtmlBlockState(trimmed, options, out var blockState)) return false;

            int j = i;
            var segments = new List<string>();
            int stackDepth = 0;

            while (j < lines.Length) {
                string current = lines[j];
                string normalized = current.Length > 0 && current[current.Length - 1] == '\r'
                    ? current.TrimEnd('\r')
                    : current;
                segments.Add(normalized);
                if (blockState.TracksStack) {
                    stackDepth = UpdateStackDepth(blockState, normalized, stackDepth);
                }
                bool completed = blockState.IsSatisfiedBy(normalized);
                j++;

                if (completed) break;
                if (blockState.TracksStack
                    && stackDepth <= 0
                    && blockState.Kind is HtmlBlockKind.Type6 or HtmlBlockKind.Type7
                    && string.Equals(blockState.PrimaryTagName, "details", StringComparison.OrdinalIgnoreCase)) {
                    break;
                }

                if (blockState.EndsOnBlankLine && j < lines.Length && string.IsNullOrWhiteSpace(lines[j])) {
                    if (!blockState.TracksStack) {
                        break;
                    }

                    if (!blockState.AllowsBlankLineContinuation || stackDepth <= 0) {
                        break;
                    }
                }
            }

            string htmlContent = string.Join("\n", segments);
            if (string.Equals(blockState.PrimaryTagName, "details", StringComparison.OrdinalIgnoreCase) &&
                TryParseDetails(htmlContent, i, options, state, out var details)) {
                doc.Add(details!);
            } else if (blockState.Kind == HtmlBlockKind.Type2) {
                doc.Add(new HtmlCommentBlock(htmlContent));
            } else {
                doc.Add(new HtmlRawBlock(htmlContent));
            }
            i = j;
            return true;
        }

        internal static bool IsParagraphInterruptingHtmlBlockStart(string line, MarkdownReaderOptions options) {
            if (options?.HtmlBlocks != true) return false;
            if (string.IsNullOrEmpty(line)) return false;

            int indent = CountLeadingSpaces(line);
            if (indent < 0 || indent > 3) return false;

            string trimmed = indent == 0 ? line : line.Substring(indent);
            if (trimmed.Length == 0 || trimmed[0] != '<') return false;
            if (TryParseAngleAutolink(trimmed, 0, out _, out _, out _)) return false;
            if (!TryGetHtmlBlockState(trimmed, options, out var blockState)) return false;

            return blockState.Kind != HtmlBlockKind.Type7;
        }

        private static bool IsHeaderlessSingleRowTableMarker(string trimmed) {
            return string.Equals(trimmed.Trim(), TableBlock.HeaderlessSingleRowTableMarker, StringComparison.Ordinal);
        }

        private static bool TryParseHeaderlessSingleRowTable(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Tables) {
                return false;
            }

            int tableStart = i + 1;
            if (tableStart >= lines.Length || !LooksLikeTableRow(lines[tableStart])) {
                return false;
            }

            if (!TryGetTableExtent(lines, tableStart, out int end, out _, allowSingleRowHeaderless: true)) {
                return false;
            }

            TableBlock table = ParseTable(lines, tableStart, end, options, state);
            if (table.Headers.Count == 0 && table.Rows.Count == 1) {
                table.PreserveHeaderlessSingleRowTable = true;
            }

            doc.Add(table);
            i = end + 1;
            return true;
        }

        private static bool TryParseDetails(string htmlContent, int startLineIndex, MarkdownReaderOptions options, MarkdownReaderState state, out DetailsBlock? block) {
            block = null;
            if (string.IsNullOrWhiteSpace(htmlContent)) return false;

            int tagEnd = htmlContent.IndexOf('>');
            if (tagEnd < 0) return false;

            int openingStart = htmlContent.IndexOf('<');
            if (openingStart < 0 || openingStart > tagEnd) return false;

            string openTag = htmlContent.Substring(0, tagEnd + 1).Trim();
            if (!openTag.StartsWith("<details", StringComparison.OrdinalIgnoreCase)) return false;

            bool isOpen = openTag.IndexOf(" open", StringComparison.OrdinalIgnoreCase) >= 0;

            int closeIdx = htmlContent.LastIndexOf("</details>", StringComparison.OrdinalIgnoreCase);
            if (closeIdx < 0 || closeIdx <= tagEnd) return false;
            int trailingStart = closeIdx + "</details>".Length;
            if (trailingStart < htmlContent.Length && !string.IsNullOrWhiteSpace(htmlContent.Substring(trailingStart))) {
                return false;
            }

            string inner = htmlContent.Substring(tagEnd + 1, closeIdx - tagEnd - 1);

            SummaryBlock? summary = null;
            int bodyStart = 0;
            int summaryStart = inner.IndexOf("<summary", StringComparison.OrdinalIgnoreCase);
            if (summaryStart >= 0) {
                int summaryTagEnd = inner.IndexOf('>', summaryStart);
                if (summaryTagEnd < 0) return false;
                int summaryClose = inner.IndexOf("</summary>", summaryTagEnd + 1, StringComparison.OrdinalIgnoreCase);
                if (summaryClose < 0) return false;

                string summarySourceText = inner.Substring(summaryTagEnd + 1, summaryClose - summaryTagEnd - 1);
                string summaryInner = summarySourceText.Trim();
                var decoded = System.Net.WebUtility.HtmlDecode(summaryInner);
                var inlines = ParseInlines(decoded, options, state);
                int summaryStartInHtml = tagEnd + 1 + summaryStart;
                int summaryTagEndInHtml = tagEnd + 1 + summaryTagEnd;
                int summaryCloseInHtml = tagEnd + 1 + summaryClose;
                int summaryCloseEndInHtml = summaryCloseInHtml + "</summary>".Length - 1;
                int summaryTextStartInHtml = tagEnd + 1 + summaryTagEnd + 1;
                int summaryTextEndInHtml = summaryCloseInHtml - 1;
                summary = new SummaryBlock(inlines) {
                    SyntaxSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryStartInHtml, summaryCloseEndInHtml),
                    OpeningTag = inner.Substring(summaryStart, summaryTagEnd - summaryStart + 1),
                    ClosingTag = inner.Substring(summaryClose, "</summary>".Length),
                    SourceText = summarySourceText,
                    OpeningTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryStartInHtml, summaryTagEndInHtml),
                    ClosingTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryCloseInHtml, summaryCloseEndInHtml)
                };

                if (summaryTextStartInHtml <= summaryTextEndInHtml) {
                    summary.TextSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, summaryTextStartInHtml, summaryTextEndInHtml);
                }

                bodyStart = summaryClose + "</summary>".Length;
            }

            string body = inner.Substring(bodyStart);
            int bodyLineOffset = state.SourceLineOffset + startLineIndex + CountNewLines(htmlContent, 0, tagEnd + 1) + CountNewLines(inner, 0, bodyStart);
            var (childBlocks, syntaxChildren) = ParseNestedMarkdownBlocks(body, options, state, bodyLineOffset);
            block = new DetailsBlock(summary, childBlocks, isOpen) {
                InsertBlankLineAfterSummary = body.StartsWith("\n\n", StringComparison.Ordinal),
                InsertBlankLineBeforeClosing = body.EndsWith("\n\n", StringComparison.Ordinal),
                OpeningTag = htmlContent.Substring(openingStart, tagEnd - openingStart + 1),
                ClosingTag = htmlContent.Substring(closeIdx, "</details>".Length),
                OpeningTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, openingStart, tagEnd),
                ClosingTagSourceSpan = CreateDetailsSourceSpan(state, htmlContent, startLineIndex, closeIdx, closeIdx + "</details>".Length - 1)
            };
            block.SyntaxChildren = syntaxChildren;
            return true;
        }

        private static MarkdownSourceSpan CreateDetailsSourceSpan(
            MarkdownReaderState state,
            string htmlContent,
            int startLineIndex,
            int startIndexInclusive,
            int endIndexInclusive) {
            GetDetailsSourcePosition(htmlContent, startIndexInclusive, out int startLineOffset, out int startColumnOffset);
            GetDetailsSourcePosition(htmlContent, endIndexInclusive, out int endLineOffset, out int endColumnOffset);

            int absoluteStartLine = state.SourceLineOffset + startLineIndex + startLineOffset + 1;
            int absoluteEndLine = state.SourceLineOffset + startLineIndex + endLineOffset + 1;

            return CreateSpan(
                state,
                absoluteStartLine,
                startColumnOffset + 1,
                absoluteEndLine,
                endColumnOffset + 1);
        }

        private static void GetDetailsSourcePosition(string htmlContent, int index, out int lineOffset, out int columnOffset) {
            int clampedIndex = Math.Max(0, Math.Min(index, Math.Max(0, htmlContent.Length - 1)));
            lineOffset = 0;
            int lineStart = 0;

            for (int i = 0; i < clampedIndex; i++) {
                if (htmlContent[i] == '\n') {
                    lineOffset++;
                    lineStart = i + 1;
                }
            }

            columnOffset = clampedIndex - lineStart;
        }

        private static int CountNewLines(string text, int start, int length) {
            if (string.IsNullOrEmpty(text) || length <= 0 || start >= text.Length) return 0;

            int end = Math.Min(text.Length, start + length);
            int count = 0;
            for (int i = start; i < end; i++) {
                if (text[i] == '\n') count++;
            }
            return count;
        }

        private static bool TryGetHtmlBlockState(string trimmedLine, MarkdownReaderOptions options, out HtmlBlockState state) {
            if (IsType1Start(trimmedLine, out var token)) {
                state = new HtmlBlockState(HtmlBlockKind.Type1, token, null, allowsBlankLineContinuation: false);
                return true;
            }

            if (trimmedLine.StartsWith("<!--", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type2, null, null, allowsBlankLineContinuation: false);
                return true;
            }

            if (trimmedLine.StartsWith("<?", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type3, null, null, allowsBlankLineContinuation: false);
                return true;
            }

            if (IsType4Start(trimmedLine)) {
                state = new HtmlBlockState(HtmlBlockKind.Type4, null, null, allowsBlankLineContinuation: false);
                return true;
            }

            if (trimmedLine.StartsWith("<![CDATA[", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type5, null, null, allowsBlankLineContinuation: false);
                return true;
            }

            if (IsType6Start(trimmedLine, options, out var tagName, out var allowsBlankLines)) {
                state = new HtmlBlockState(HtmlBlockKind.Type6, null, tagName, allowsBlankLines);
                return true;
            }

            if (IsType7Start(trimmedLine, options, out var type7Tag, out var allowsBlankLinesType7)) {
                state = new HtmlBlockState(HtmlBlockKind.Type7, null, type7Tag, allowsBlankLinesType7);
                return true;
            }

            state = default;
            return false;
        }

        private static bool IsType1Start(string trimmedLine, out string? endToken) {
            if (!TryReadHtmlTagNamePrefix(trimmedLine, out var tagName, out var isClosing, out _, allowLineEnd: true)) {
                endToken = null;
                return false;
            }

            if (isClosing) {
                endToken = null;
                return false;
            }

            string[] tags = { "script", "pre", "style", "textarea" };
            foreach (var tag in tags) {
                if (string.Equals(tagName, tag, StringComparison.OrdinalIgnoreCase)) {
                    endToken = $"</{tag}>";
                    return true;
                }
            }

            endToken = null;
            return false;
        }

        private static bool IsType4Start(string trimmedLine) {
            if (!trimmedLine.StartsWith("<!", StringComparison.Ordinal) || trimmedLine.Length < 3) return false;
            char next = trimmedLine[2];
            return next is >= 'A' and <= 'Z';
        }

        private static bool IsType6Start(string trimmedLine, MarkdownReaderOptions options, out string? tagName, out bool allowsBlankLines) {
            allowsBlankLines = false;
            tagName = null;

            string? parsedName;
            bool isClosing;
            bool parsed = options.AllowLooseHtmlBlockStartTags
                ? TryReadHtmlTagNamePrefix(trimmedLine, out parsedName, out isClosing, out _)
                : TryParseTag(trimmedLine, out parsedName, out isClosing, out var endIndex) && endIndex >= 0;
            if (!parsed) return false;
            if (!s_BlockTags.Contains(parsedName!)) return false;

            if (!isClosing) {
                allowsBlankLines = AllowsBlankLineContinuation(parsedName!, options);
            }

            tagName = parsedName;
            return true;
        }

        private static bool IsType7Start(string trimmedLine, MarkdownReaderOptions options, out string? tagName, out bool allowsBlankLines) {
            allowsBlankLines = false;
            if (!TryParseTag(trimmedLine, out tagName, out var isClosing, out var endIndex)) return false;
            if (endIndex < 0) return false;
            if (s_BlockTags.Contains(tagName!)) return false;
            if (!isClosing && string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "pre", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "textarea", StringComparison.OrdinalIgnoreCase)) return false;
            if (!IsOnlyWhitespaceAfter(trimmedLine, endIndex + 1)) return false;
            allowsBlankLines = AllowsBlankLineContinuation(tagName!, options);
            return true;
        }

        private static bool IsOnlyWhitespaceAfter(string line, int startIndex) {
            for (int i = startIndex; i < line.Length; i++) {
                if (!IsHtmlAttributeWhitespace(line[i])) {
                    return false;
                }
            }

            return true;
        }

        private static bool AllowsBlankLineContinuation(string tagName, MarkdownReaderOptions options) {
            return options.PreserveHtmlBlockBlankLineContent && s_BlankLineFriendlyTags.Contains(tagName);
        }

        internal static bool IsBlockOrRawTextHtmlTagName(string tagName) {
            if (string.IsNullOrEmpty(tagName)) {
                return false;
            }

            return s_BlockTags.Contains(tagName)
                || string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "pre", StringComparison.OrdinalIgnoreCase)
                || string.Equals(tagName, "textarea", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool TryParseHtmlTag(string line, out string tagName, out bool isClosing, out int endIndex) =>
            TryParseTag(line, out tagName, out isClosing, out endIndex);

        private static readonly HashSet<string> s_BlankLineFriendlyTags = new(StringComparer.OrdinalIgnoreCase) {
            "details",
            "table",
        };

        private static int CountLeadingSpaces(string line) {
            int count = 0;
            while (count < line.Length && line[count] == ' ') {
                count++;
            }

            return count;
        }

        private static bool TryParseTag(string line, out string tagName, out bool isClosing, out int endIndex) {
            tagName = string.Empty;
            isClosing = false;
            endIndex = -1;

            if (line.Length < 2 || line[0] != '<') return false;

            int idx = 1;
            if (idx < line.Length && line[idx] == '/') {
                isClosing = true;
                idx++;
            }

            if (idx >= line.Length || !IsAsciiLetter(line[idx])) return false;
            int nameStart = idx;
            idx++;
            while (idx < line.Length && IsHtmlTagNameContinuation(line[idx])) idx++;
            tagName = line.Substring(nameStart, idx - nameStart);
            if (tagName.Length == 0) return false;

            return TryConsumeHtmlTagRemainder(line, idx, isClosing, out endIndex);
        }

        private static bool TryReadHtmlTagNamePrefix(string line, out string tagName, out bool isClosing, out int nameEndIndex, bool allowLineEnd = false) {
            tagName = string.Empty;
            isClosing = false;
            nameEndIndex = -1;

            if (line.Length < 2 || line[0] != '<') return false;

            int idx = 1;
            if (idx < line.Length && line[idx] == '/') {
                isClosing = true;
                idx++;
            }

            if (idx >= line.Length || !IsAsciiLetter(line[idx])) return false;
            int nameStart = idx;
            idx++;
            while (idx < line.Length && IsHtmlTagNameContinuation(line[idx])) idx++;
            tagName = line.Substring(nameStart, idx - nameStart);
            nameEndIndex = idx - 1;

            if (idx >= line.Length) return allowLineEnd;

            char next = line[idx];
            if (next == '>') return true;
            if (next == '/' && idx + 1 < line.Length && line[idx + 1] == '>') return true;
            return IsHtmlAttributeWhitespace(next);
        }

        private static bool TryConsumeHtmlTagRemainder(string line, int index, bool isClosing, out int endIndex) {
            endIndex = -1;
            if (index < 0 || index >= line.Length) return false;

            int idx = index;
            if (isClosing) {
                while (idx < line.Length && IsHtmlAttributeWhitespace(line[idx])) idx++;
                if (idx < line.Length && line[idx] == '>') {
                    endIndex = idx;
                    return true;
                }

                return false;
            }

            while (idx < line.Length) {
                bool sawAttributeWhitespace = false;
                while (idx < line.Length && IsHtmlAttributeWhitespace(line[idx])) idx++;
                sawAttributeWhitespace = idx > index;
                if (idx >= line.Length) return false;

                if (line[idx] == '>') {
                    endIndex = idx;
                    return true;
                }

                if (line[idx] == '/' && idx + 1 < line.Length && line[idx + 1] == '>') {
                    endIndex = idx + 1;
                    return true;
                }

                if (!sawAttributeWhitespace) return false;
                if (!TryConsumeHtmlAttribute(line, ref idx)) return false;
                index = idx;
            }

            return false;
        }

        private static bool TryConsumeHtmlAttribute(string line, ref int index) {
            if (index < 0 || index >= line.Length || !IsHtmlAttributeNameStart(line[index])) return false;

            index++;
            while (index < line.Length && IsHtmlAttributeNameContinuation(line[index])) index++;

            int afterName = index;
            while (index < line.Length && IsHtmlAttributeWhitespace(line[index])) index++;
            if (index >= line.Length || line[index] != '=') {
                index = afterName;
                return true;
            }

            index++;
            while (index < line.Length && IsHtmlAttributeWhitespace(line[index])) index++;
            if (index >= line.Length) return false;

            char quote = line[index];
            if (quote == '"' || quote == '\'') {
                index++;
                while (index < line.Length && line[index] != quote) {
                    if (line[index] == '\\' && index + 1 < line.Length) {
                        index += 2;
                        continue;
                    }

                    index++;
                }

                if (index >= line.Length) return false;
                index++;
                return true;
            }

            int valueStart = index;
            while (index < line.Length && !IsHtmlAttributeWhitespace(line[index]) && line[index] != '>') {
                char ch = line[index];
                if (ch == '"' || ch == '\'' || ch == '=' || ch == '<' || ch == '`') return false;
                index++;
            }

            return index > valueStart;
        }

        private static bool IsHtmlTagNameContinuation(char value) =>
            IsAsciiLetter(value) || char.IsDigit(value) || value == '-';

        private static bool IsHtmlAttributeNameStart(char value) =>
            IsAsciiLetter(value) || value == '_' || value == ':';

        private static bool IsHtmlAttributeNameContinuation(char value) =>
            IsHtmlAttributeNameStart(value) || char.IsDigit(value) || value == '.' || value == '-';

        private static int UpdateStackDepth(HtmlBlockState state, string line, int currentDepth) {
            if (string.IsNullOrEmpty(state.PrimaryTagName)) return currentDepth;

            string tagName = state.PrimaryTagName!;
            int index = 0;

            while (index < line.Length) {
                int tagStart = line.IndexOf('<', index);
                if (tagStart < 0) break;

                if (!TryParseTagAt(line, tagStart, out var foundTag, out var isClosing, out var nextIndex, out var endIndex)) {
                    index = tagStart + 1;
                    continue;
                }

                bool isTarget = string.Equals(foundTag, tagName, StringComparison.OrdinalIgnoreCase);
                bool selfClosing = IsSelfClosing(line, tagStart, endIndex);

                if (isTarget && !selfClosing) {
                    currentDepth += isClosing ? -1 : 1;
                }

                index = nextIndex;
            }

            return currentDepth;
        }

        private static bool TryParseTagAt(string line, int startIndex, out string tagName, out bool isClosing, out int nextIndex, out int endIndex) {
            tagName = string.Empty;
            isClosing = false;
            nextIndex = startIndex + 1;
            endIndex = -1;

            if (startIndex < 0 || startIndex >= line.Length || line[startIndex] != '<') return false;

            string segment = line.Substring(startIndex);
            if (!TryParseTag(segment, out tagName, out isClosing, out var localEndIndex)) {
                return false;
            }

            endIndex = startIndex + localEndIndex;
            nextIndex = endIndex + 1;
            return true;
        }

        private static bool IsSelfClosing(string line, int startIndex, int endIndex) {
            if (endIndex <= startIndex) return false;

            int idx = endIndex - 1;
            while (idx > startIndex && char.IsWhiteSpace(line[idx])) idx--;
            if (idx <= startIndex || line[idx] != '/') return false;

            int previous = idx - 1;
            while (previous > startIndex && char.IsWhiteSpace(line[previous])) previous--;

            return previous >= startIndex && line[previous] != '/';
        }
    }
}
