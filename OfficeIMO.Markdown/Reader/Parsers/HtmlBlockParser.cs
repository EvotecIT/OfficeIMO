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
            public HtmlBlockState(HtmlBlockKind kind, string? primaryEndToken) {
                Kind = kind;
                PrimaryEndToken = primaryEndToken;
            }

            public HtmlBlockKind Kind { get; }
            public string? PrimaryEndToken { get; }

            public bool EndsOnBlankLine => Kind is HtmlBlockKind.Type6 or HtmlBlockKind.Type7;

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
            if (lines == null || lines.Length == 0) return false;

            var line = lines[i];
            if (string.IsNullOrEmpty(line)) return false;

            int indent = CountLeadingSpaces(line);
            if (indent < 0 || indent > 3) return false;

            string trimmed = indent == 0 ? line : line.Substring(indent);
            if (trimmed.Length == 0 || trimmed[0] != '<') return false;

            if (!TryGetHtmlBlockState(trimmed, out var blockState)) return false;

            int j = i;
            var segments = new List<string>();

            while (j < lines.Length) {
                string current = lines[j];
                string normalized = current.EndsWith('\r') ? current.TrimEnd('\r') : current;
                segments.Add(normalized);
                bool completed = blockState.IsSatisfiedBy(normalized);
                j++;

                if (completed) break;
                if (blockState.EndsOnBlankLine && j < lines.Length && string.IsNullOrWhiteSpace(lines[j])) break;
            }

            doc.Add(new HtmlRawBlock(string.Join('\n', segments)));
            i = j;
            return true;
        }

        private static bool TryGetHtmlBlockState(string trimmedLine, out HtmlBlockState state) {
            if (IsType1Start(trimmedLine, out var token)) {
                state = new HtmlBlockState(HtmlBlockKind.Type1, token);
                return true;
            }

            if (trimmedLine.StartsWith("<!--", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type2, null);
                return true;
            }

            if (trimmedLine.StartsWith("<?", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type3, null);
                return true;
            }

            if (IsType4Start(trimmedLine)) {
                state = new HtmlBlockState(HtmlBlockKind.Type4, null);
                return true;
            }

            if (trimmedLine.StartsWith("<![CDATA[", StringComparison.Ordinal)) {
                state = new HtmlBlockState(HtmlBlockKind.Type5, null);
                return true;
            }

            if (IsType6Start(trimmedLine)) {
                state = new HtmlBlockState(HtmlBlockKind.Type6, null);
                return true;
            }

            if (IsType7Start(trimmedLine)) {
                state = new HtmlBlockState(HtmlBlockKind.Type7, null);
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

        private static bool IsType6Start(string trimmedLine) {
            if (!TryParseTag(trimmedLine, out var tagName, out _, out _)) return false;
            return s_BlockTags.Contains(tagName);
        }

        private static bool IsType7Start(string trimmedLine) {
            if (!TryParseTag(trimmedLine, out var tagName, out var isClosing, out _)) return false;
            if (s_BlockTags.Contains(tagName)) return false;
            if (!isClosing && string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "pre", StringComparison.OrdinalIgnoreCase)) return false;
            return true;
        }

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

            for (int j = idx; j < line.Length; j++) {
                if (!char.IsWhiteSpace(line[j])) return false;
            }

            return true;
        }
    }
}
