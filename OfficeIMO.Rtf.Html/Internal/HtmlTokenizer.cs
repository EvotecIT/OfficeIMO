namespace OfficeIMO.Rtf.Html;

internal static partial class HtmlTokenizer {
    internal static IReadOnlyList<HtmlToken> Tokenize(string html, RtfHtmlReadOptions? options = null) {
        var tokens = new List<HtmlToken>();
        HtmlTokenLimitTracker? limits = HtmlTokenLimitTracker.Create(options);
        int index = 0;
        while (index < html.Length) {
            int tagStart = html.IndexOf('<', index);
            if (tagStart < 0) {
                AddText(tokens, html.Substring(index), limits);
                break;
            }

            if (tagStart > index) {
                AddText(tokens, html.Substring(index, tagStart - index), limits);
            }

            if (StartsWith(html, tagStart, "<!--")) {
                int commentEnd = html.IndexOf("-->", tagStart + 4, StringComparison.Ordinal);
                index = commentEnd < 0 ? html.Length : commentEnd + 3;
                continue;
            }

            int tagEnd = FindTagEnd(html, tagStart + 1);
            if (tagEnd < 0) {
                AddText(tokens, html.Substring(tagStart), limits);
                break;
            }

            AddTag(tokens, html.Substring(tagStart + 1, tagEnd - tagStart - 1), limits);
            index = tagEnd + 1;
        }

        return tokens;
    }

    private static void AddText(List<HtmlToken> tokens, string text, HtmlTokenLimitTracker? limits) {
        if (text.Length > 0) {
            limits?.RecordText();
            tokens.Add(new HtmlToken(HtmlTokenKind.Text, WebUtility.HtmlDecode(text)));
        }
    }

    private static void AddTag(List<HtmlToken> tokens, string rawTag, HtmlTokenLimitTracker? limits) {
        string tag = rawTag.Trim();
        if (tag.Length == 0 || tag[0] == '!' || tag[0] == '?') {
            return;
        }

        if (tag[0] == '/') {
            string name = ReadName(tag, 1).ToLowerInvariant();
            if (name.Length > 0) {
                limits?.RecordEnd(name);
                tokens.Add(new HtmlToken(HtmlTokenKind.EndTag, name));
            }

            return;
        }

        bool selfClosing = tag.EndsWith("/", StringComparison.Ordinal);
        if (selfClosing) {
            tag = tag.Substring(0, tag.Length - 1).TrimEnd();
        }

        string startName = ReadName(tag, 0).ToLowerInvariant();
        if (startName.Length == 0) {
            return;
        }

        bool isSelfClosing = selfClosing || IsVoidElement(startName);
        limits?.RecordStart(startName, isSelfClosing);
        tokens.Add(new HtmlToken(HtmlTokenKind.StartTag, startName, ParseAttributes(tag, startName.Length), isSelfClosing));
    }

    private static Dictionary<string, string> ParseAttributes(string tag, int index) {
        var attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        while (index < tag.Length) {
            SkipWhitespace(tag, ref index);
            string name = ReadName(tag, index);
            if (name.Length == 0) {
                break;
            }

            index += name.Length;
            SkipWhitespace(tag, ref index);
            string value = string.Empty;
            if (index < tag.Length && tag[index] == '=') {
                index++;
                SkipWhitespace(tag, ref index);
                value = ReadAttributeValue(tag, ref index);
            }

            attributes[name.ToLowerInvariant()] = WebUtility.HtmlDecode(value);
        }

        return attributes;
    }

    private static string ReadAttributeValue(string tag, ref int index) {
        if (index >= tag.Length) {
            return string.Empty;
        }

        char quote = tag[index];
        if (quote == '"' || quote == '\'') {
            int start = ++index;
            while (index < tag.Length && tag[index] != quote) {
                index++;
            }

            string value = tag.Substring(start, index - start);
            if (index < tag.Length) {
                index++;
            }

            return value;
        }

        int unquotedStart = index;
        while (index < tag.Length && !char.IsWhiteSpace(tag[index])) {
            index++;
        }

        return tag.Substring(unquotedStart, index - unquotedStart);
    }

    private static string ReadName(string text, int index) {
        int start = index;
        while (index < text.Length && (char.IsLetterOrDigit(text[index]) || text[index] == '-' || text[index] == ':' || text[index] == '_')) {
            index++;
        }

        return text.Substring(start, index - start);
    }

    private static int FindTagEnd(string html, int index) {
        char quote = '\0';
        for (int i = index; i < html.Length; i++) {
            char current = html[i];
            if (quote != '\0') {
                if (current == quote) {
                    quote = '\0';
                }
            } else if (current == '"' || current == '\'') {
                quote = current;
            } else if (current == '>') {
                return i;
            }
        }

        return -1;
    }

    private static void SkipWhitespace(string text, ref int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }
    }

    private static bool StartsWith(string text, int index, string value) {
        return index + value.Length <= text.Length && string.CompareOrdinal(text, index, value, 0, value.Length) == 0;
    }

    private static bool IsVoidElement(string name) {
        switch (name) {
            case "area":
            case "base":
            case "br":
            case "col":
            case "embed":
            case "hr":
            case "img":
            case "input":
            case "link":
            case "meta":
            case "param":
            case "source":
            case "track":
            case "wbr":
                return true;
            default:
                return false;
        }
    }
}
