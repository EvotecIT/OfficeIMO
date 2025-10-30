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
            var sb = new System.Text.StringBuilder();

            while (j < lines.Length) {
                string current = lines[j];
                sb.AppendLine(current);
                bool completed = blockState.IsSatisfiedBy(current);
                j++;

                if (completed) break;
                if (blockState.EndsOnBlankLine && j < lines.Length && string.IsNullOrWhiteSpace(lines[j])) break;
            }

            doc.Add(new HtmlRawBlock(sb.ToString().TrimEnd('\r', '\n')));
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
            string[] tags = { "script", "pre", "style" };
            foreach (var tag in tags) {
                if (StartsWithTag(trimmedLine, tag)) {
                    endToken = $"</{tag}>";
                    return true;
                }
            }

            endToken = null;
            return false;
        }

        private static bool StartsWithTag(string line, string tag) {
            if (!line.StartsWith("<", StringComparison.Ordinal)) return false;

            int idx = 1;
            int len = line.Length;

            while (idx < len && char.IsWhiteSpace(line[idx])) idx++;

            if (idx + tag.Length > len) return false;
            if (string.Compare(line, idx, tag, 0, tag.Length, StringComparison.OrdinalIgnoreCase) != 0) return false;
            idx += tag.Length;

            if (idx >= len) return true;
            char c = line[idx];
            return char.IsWhiteSpace(c) || c == '>' || c == '/';
        }

        private static bool IsType4Start(string trimmedLine) {
            if (!trimmedLine.StartsWith("<!", StringComparison.Ordinal) || trimmedLine.Length < 3) return false;
            char next = trimmedLine[2];
            return next is >= 'A' and <= 'Z';
        }

        private static bool IsType6Start(string trimmedLine) {
            if (trimmedLine.Length < 2 || trimmedLine[0] != '<') return false;

            int idx = 1;
            if (idx < trimmedLine.Length && trimmedLine[idx] == '/') idx++;

            int nameStart = idx;
            if (idx >= trimmedLine.Length || !char.IsLetter(trimmedLine[idx])) return false;
            idx++;
            while (idx < trimmedLine.Length && (char.IsLetterOrDigit(trimmedLine[idx]) || trimmedLine[idx] == '-')) idx++;
            string tagName = trimmedLine.Substring(nameStart, idx - nameStart);
            if (!s_BlockTags.Contains(tagName)) return false;

            if (idx >= trimmedLine.Length) return false;

            bool insideQuotes = false;
            char quoteChar = '\0';
            int end = -1;

            while (idx < trimmedLine.Length) {
                char current = trimmedLine[idx];

                if (insideQuotes) {
                    if (current == quoteChar) {
                        insideQuotes = false;
                        quoteChar = '\0';
                    }
                } else {
                    if (current == '\'' || current == '"') {
                        insideQuotes = true;
                        quoteChar = current;
                    } else if (current == '>') {
                        end = idx;
                        idx++;
                        break;
                    } else if (current == '<') {
                        return false;
                    }
                }

                idx++;
            }

            if (insideQuotes || end < 0) return false;
            return true;
        }

        private static bool IsType7Start(string trimmedLine) {
            if (trimmedLine.Length < 3 || trimmedLine[0] != '<') return false;

            bool isClosing = trimmedLine[1] == '/';
            int idx = isClosing ? 2 : 1;
            if (idx >= trimmedLine.Length) return false;

            int nameStart = idx;
            if (!char.IsLetter(trimmedLine[idx])) return false;
            idx++;
            while (idx < trimmedLine.Length && (char.IsLetterOrDigit(trimmedLine[idx]) || trimmedLine[idx] == '-')) idx++;
            string tagName = trimmedLine.Substring(nameStart, idx - nameStart);
            if (!isClosing && string.Equals(tagName, "script", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "style", StringComparison.OrdinalIgnoreCase)) return false;
            if (!isClosing && string.Equals(tagName, "pre", StringComparison.OrdinalIgnoreCase)) return false;

            int end = trimmedLine.IndexOf('>');
            if (end < 0) return false;

            if (trimmedLine[end] == '>') {
                for (int j = end + 1; j < trimmedLine.Length; j++) {
                    if (!char.IsWhiteSpace(trimmedLine[j])) return false;
                }
                return true;
            }

            return false;
        }

        private static int CountLeadingSpaces(string line) {
            int count = 0;
            while (count < line.Length && line[count] == ' ') {
                count++;
            }

            return count;
        }
    }
}
