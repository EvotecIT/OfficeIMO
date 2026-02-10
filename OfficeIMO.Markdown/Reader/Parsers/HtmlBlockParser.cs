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

            // Avoid treating angle-bracket autolinks like "<https://...>" as HTML blocks.
            if (TryParseAngleAutolink(trimmed, 0, out _, out _, out _)) return false;

            if (!TryGetHtmlBlockState(trimmed, out var blockState)) return false;

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
                if (blockState.TracksStack && stackDepth <= 0 && blockState.Kind is HtmlBlockKind.Type6 or HtmlBlockKind.Type7) break;
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
                TryParseDetails(htmlContent, options, state, out var details)) {
                doc.Add(details!);
            } else if (blockState.Kind == HtmlBlockKind.Type2) {
                doc.Add(new HtmlCommentBlock(htmlContent));
            } else {
                doc.Add(new HtmlRawBlock(htmlContent));
            }
            i = j;
            return true;
        }

        private static bool TryParseDetails(string htmlContent, MarkdownReaderOptions options, MarkdownReaderState state, out DetailsBlock? block) {
            block = null;
            if (string.IsNullOrWhiteSpace(htmlContent)) return false;

            int tagEnd = htmlContent.IndexOf('>');
            if (tagEnd < 0) return false;

            string openTag = htmlContent.Substring(0, tagEnd + 1).Trim();
            if (!openTag.StartsWith("<details", StringComparison.OrdinalIgnoreCase)) return false;

            bool isOpen = openTag.IndexOf(" open", StringComparison.OrdinalIgnoreCase) >= 0;

            int closeIdx = htmlContent.LastIndexOf("</details>", StringComparison.OrdinalIgnoreCase);
            if (closeIdx < 0 || closeIdx <= tagEnd) return false;

            string inner = htmlContent.Substring(tagEnd + 1, closeIdx - tagEnd - 1);

            SummaryBlock? summary = null;
            int bodyStart = 0;
            int summaryStart = inner.IndexOf("<summary", StringComparison.OrdinalIgnoreCase);
            if (summaryStart >= 0) {
                int summaryTagEnd = inner.IndexOf('>', summaryStart);
                if (summaryTagEnd < 0) return false;
                int summaryClose = inner.IndexOf("</summary>", summaryTagEnd + 1, StringComparison.OrdinalIgnoreCase);
                if (summaryClose < 0) return false;

                string summaryInner = inner.Substring(summaryTagEnd + 1, summaryClose - summaryTagEnd - 1).Trim();
                var decoded = System.Net.WebUtility.HtmlDecode(summaryInner);
                var inlines = ParseInlines(decoded, options, state);
                summary = new SummaryBlock(inlines);
                bodyStart = summaryClose + "</summary>".Length;
            }

            string body = inner.Substring(bodyStart);
            var nestedOptions = CloneOptionsWithoutFrontMatter(options);
            var nestedState = CloneState(state);
            var nestedDoc = ParseInternal(body, nestedOptions, nestedState, allowFrontMatter: false);
            block = new DetailsBlock(summary, nestedDoc.Blocks, isOpen) {
                InsertBlankLineAfterSummary = body.StartsWith("\n\n", StringComparison.Ordinal),
                InsertBlankLineBeforeClosing = body.EndsWith("\n\n", StringComparison.Ordinal)
            };
            return true;
        }

        private static bool TryGetHtmlBlockState(string trimmedLine, out HtmlBlockState state) {
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

            if (IsType6Start(trimmedLine, out var tagName, out var allowsBlankLines)) {
                state = new HtmlBlockState(HtmlBlockKind.Type6, null, tagName, allowsBlankLines);
                return true;
            }

            if (IsType7Start(trimmedLine, out var type7Tag, out var allowsBlankLinesType7)) {
                state = new HtmlBlockState(HtmlBlockKind.Type7, null, type7Tag, allowsBlankLinesType7);
                return true;
            }

            state = default;
            return false;
        }

        private static bool IsType1Start(string trimmedLine, out string? endToken) {
            if (!TryParseTag(trimmedLine, out var tagName, out var isClosing, out _)) {
                endToken = null;
                return false;
            }

            if (isClosing) {
                endToken = null;
                return false;
            }

            string[] tags = { "script", "pre", "style" };
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

        private static bool IsType6Start(string trimmedLine, out string? tagName, out bool allowsBlankLines) {
            allowsBlankLines = false;
            tagName = null;

            if (!TryParseTag(trimmedLine, out var parsedName, out var isClosing, out var endIndex)) return false;
            if (endIndex < 0) return false;
            if (!s_BlockTags.Contains(parsedName!)) return false;

            if (!isClosing) {
                allowsBlankLines = AllowsBlankLineContinuation(parsedName!);
            }

            tagName = parsedName;
            return true;
        }

        private static bool IsType7Start(string trimmedLine, out string? tagName, out bool allowsBlankLines) {
            allowsBlankLines = false;
            if (!TryParseTag(trimmedLine, out tagName, out var isClosing, out var endIndex)) return false;
            if (endIndex < 0) return false;
            if (s_BlockTags.Contains(tagName!)) return false;
            if (!isClosing && string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "pre", StringComparison.OrdinalIgnoreCase)) return false;
            allowsBlankLines = AllowsBlankLineContinuation(tagName!);
            return true;
        }

        private static bool AllowsBlankLineContinuation(string tagName) {
            return s_BlankLineFriendlyTags.Contains(tagName);
        }

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

            if (idx >= line.Length || !char.IsLetter(line[idx])) return false;
            int nameStart = idx;
            idx++;
            while (idx < line.Length && (char.IsLetterOrDigit(line[idx]) || line[idx] == '-' || line[idx] == ':')) idx++;
            tagName = line.Substring(nameStart, idx - nameStart);
            if (tagName.Length == 0) return false;

            bool insideQuotes = false;
            char quoteChar = '\0';
            bool escaped = false;

            while (idx < line.Length) {
                char current = line[idx];

                if (insideQuotes) {
                    if (escaped) {
                        escaped = false;
                    } else if (current == '\\') {
                        escaped = true;
                    } else if (current == quoteChar) {
                        insideQuotes = false;
                        quoteChar = '\0';
                    }
                } else {
                    if (current == '\'' || current == '"') {
                        insideQuotes = true;
                        quoteChar = current;
                    } else if (current == '>') {
                        endIndex = idx;
                        idx++;
                        break;
                    } else if (current == '<') {
                        return false;
                    }
                }

                idx++;
            }

            if (insideQuotes || endIndex < 0) return false;

            return true;
        }

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
