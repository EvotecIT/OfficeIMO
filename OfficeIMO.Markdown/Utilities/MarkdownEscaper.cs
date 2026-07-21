namespace OfficeIMO.Markdown;

using System;
using System.Text;

/// <summary>
/// Shared helper for escaping Markdown-reserved characters across inline renderers.
/// </summary>
public static class MarkdownEscaper {
    // Inline HTML is allowed in rendered Markdown (e.g., <u> for underline), so we intentionally
    // do not escape angle brackets here to preserve legitimate HTML passthroughs.
    // Allows: "Use <u>underline</u> for emphasis" -> "Use <u>underline</u> for emphasis"
    // Escapes: "Text [link](url)" -> "Text \\[link\\]\(url\)"
    private static readonly char[] GeneralReserved = ['\\', '[', ']', '(', ')', '|', '*', '_'];
    private static readonly char[] HighlightReserved = ['\\', '[', ']', '(', ')', '|', '*', '_', '='];
    private static readonly char[] UrlReserved = ['\\', '(', ')', '[', ']', '|'];

    /// <summary>Escapes Markdown-reserved punctuation in literal inline text.</summary>
    public static string EscapeText(string? text) => Escape(text, GeneralReserved);

    /// <summary>
    /// Escapes literal inline text and any line-leading syntax that Markdown would otherwise parse
    /// as a heading, quote, list, fence, thematic break, definition, or HTML block.
    /// </summary>
    public static string EscapeTextAndLineStarts(string? text) => EscapeMarkdownLineStarts(EscapeText(text));

    /// <summary>
    /// Escapes literal text so Markdown punctuation, HTML-like text, and character-reference-like
    /// text are preserved as text when parsed again.
    /// </summary>
    public static string EscapeLiteralText(string? text) => EscapeMarkdownLineStarts(EncodeLiteralMarkdownText(text));

    internal static string EscapeRenderedLineStarts(string? text) => string.IsNullOrEmpty(text) ? string.Empty : EscapeMarkdownLineStarts(text!);
    internal static string EscapeRenderedListItemLineStarts(string? text) => string.IsNullOrEmpty(text) ? string.Empty : EscapeMarkdownLineStarts(text!, preserveDefinitionText: true);
    internal static string EscapeLiteralTableCellText(string? text) => EncodeLiteralMarkdownText(text, encodeEntityLikeAmpersands: false);
    internal static string EscapeEmphasis(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeHighlightText(string? text) => Escape(text, HighlightReserved);
    internal static string EscapeInsertedText(string? text) => Escape(text, ['\\', '[', ']', '(', ')', '|', '*', '_', '+']);
    internal static string EscapeSuperscriptText(string? text) => Escape(text, ['\\', '[', ']', '(', ')', '|', '*', '_', '^']);
    internal static string EscapeSubscriptText(string? text) => Escape(text, ['\\', '[', ']', '(', ')', '|', '*', '_', '~']);
    internal static string EscapeLinkText(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeLinkUrl(string? text) => Escape(text, UrlReserved);
    internal static string EscapeImageAlt(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeImageSrc(string? text) => Escape(text, UrlReserved);

    /// <summary>Formats an optional Markdown link or image title using a non-colliding delimiter.</summary>
    public static string FormatOptionalTitle(string? title) {
        if (string.IsNullOrEmpty(title)) return string.Empty;

        // Titles cannot contain line breaks in Markdown link/image syntax; normalize them to spaces.
        var t = title!.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
        t = EscapeText(t);

        // Prefer a delimiter that doesn't occur in the title, to avoid having to escape it.
        if (t.IndexOf('"') < 0) return " \"" + EscapeTitleContent(t, '"') + "\"";
        if (t.IndexOf('\'') < 0) return " '" + EscapeTitleContent(t, '\'') + "'";
        if (t.IndexOf(')') < 0) return " (" + EscapeTitleContent(t, ')') + ")";

        // Fallback: escape double quotes.
        return " \"" + EscapeTitleContent(t, '"') + "\"";
    }

    private static string Escape(string? text, char[] reserved) {
        if (string.IsNullOrEmpty(text)) return string.Empty;

        StringBuilder sb = new(text!.Length);
        foreach (char c in text) {
            if (Array.IndexOf(reserved, c) >= 0) sb.Append('\\');
            sb.Append(c);
        }
        return sb.ToString();
    }

    private static string EscapeMarkdownLineStarts(string text, bool preserveDefinitionText = false) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        StringBuilder sb = new StringBuilder(text.Length + 8);
        int start = 0;
        while (start < text.Length) {
            int newlineIndex = text.IndexOf('\n', start);
            int length = newlineIndex < 0 ? text.Length - start : newlineIndex - start;
            sb.Append(EscapeMarkdownLineStart(
                text.Substring(start, length),
                preserveDefinitionText: preserveDefinitionText));
            if (newlineIndex < 0) {
                break;
            }

            sb.Append('\n');
            start = newlineIndex + 1;
        }

        return sb.ToString();
    }

    private static string EscapeMarkdownLineStart(string line, bool preserveDefinitionText = false) {
        if (line.Length == 0) {
            return line;
        }

        int markerIndex = 0;
        while (markerIndex < line.Length && markerIndex < 3 && line[markerIndex] == ' ') {
            markerIndex++;
        }

        if (markerIndex >= line.Length) {
            return line;
        }

        char marker = line[markerIndex];
        if (marker == '>'
            || IsHeadingMarker(line, markerIndex)
            || IsUnorderedListMarker(line, markerIndex)
            || IsFencedCodeMarker(line, markerIndex)
            || IsSetextHeadingUnderline(line, markerIndex)
            || IsHtmlBlockOpener(line, markerIndex)) {
            return line.Insert(markerIndex, "\\");
        }

        int orderedSeparatorIndex = GetOrderedListSeparatorIndex(line, markerIndex);
        if (orderedSeparatorIndex >= 0) {
            return line.Insert(orderedSeparatorIndex, "\\");
        }

        int definitionSeparatorIndex = GetDefinitionListSeparatorIndex(line, markerIndex);
        return definitionSeparatorIndex >= 0 && (!preserveDefinitionText || StartsWithReferenceDefinitionLikeLabel(line, markerIndex, definitionSeparatorIndex))
            ? line.Substring(0, definitionSeparatorIndex) + "&#58;" + line.Substring(definitionSeparatorIndex + 1)
            : line;
    }

    private static bool IsHeadingMarker(string line, int markerIndex) {
        int index = markerIndex;
        while (index < line.Length && line[index] == '#') {
            index++;
        }

        int markerLength = index - markerIndex;
        return markerLength is >= 1 and <= 6
               && (index >= line.Length || char.IsWhiteSpace(line[index]));
    }

    private static bool IsUnorderedListMarker(string line, int markerIndex) {
        char marker = line[markerIndex];
        if (marker != '-' && marker != '+' && marker != '*') {
            return false;
        }

        int next = markerIndex + 1;
        if (next >= line.Length || char.IsWhiteSpace(line[next])) {
            return true;
        }

        if (marker != '-') {
            return false;
        }

        int count = 0;
        for (int i = markerIndex; i < line.Length; i++) {
            if (line[i] != '-') {
                return false;
            }

            count++;
        }

        return count >= 3;
    }

    private static bool IsSetextHeadingUnderline(string line, int markerIndex) {
        if (line[markerIndex] != '=') {
            return false;
        }

        int index = markerIndex;
        while (index < line.Length && line[index] == '=') {
            index++;
        }

        while (index < line.Length) {
            if (!char.IsWhiteSpace(line[index])) {
                return false;
            }

            index++;
        }

        return index > markerIndex;
    }

    private static bool IsFencedCodeMarker(string line, int markerIndex) {
        char marker = line[markerIndex];
        if (marker != '`' && marker != '~') {
            return false;
        }

        int count = 0;
        for (int i = markerIndex; i < line.Length && line[i] == marker; i++) {
            count++;
        }

        return count >= 3;
    }

    private static bool IsHtmlBlockOpener(string line, int markerIndex) {
        if (line[markerIndex] != '<') {
            return false;
        }

        if (StartsWithAt(line, markerIndex, "<!--")) {
            return true;
        }

        if (StartsWithAt(line, markerIndex, "<?")
            || StartsWithAt(line, markerIndex, "<![CDATA[")
            || StartsWithAt(line, markerIndex, "<!\\[CDATA\\[")) {
            return true;
        }

        if (StartsWithAt(line, markerIndex, "<!")) {
            return markerIndex + 2 < line.Length && line[markerIndex + 2] >= 'A' && line[markerIndex + 2] <= 'Z';
        }

        if (StartsWithAt(line, markerIndex, "</")) {
            return markerIndex + 2 < line.Length && IsHtmlTagNameStart(line[markerIndex + 2]);
        }

        return markerIndex + 1 < line.Length && IsHtmlTagNameStart(line[markerIndex + 1]);
    }

    private static bool StartsWithAt(string text, int startIndex, string value) {
        if (startIndex < 0 || startIndex + value.Length > text.Length) {
            return false;
        }

        return string.Compare(text, startIndex, value, 0, value.Length, StringComparison.Ordinal) == 0;
    }

    private static bool IsHtmlTagNameStart(char value) {
        return value >= 'A' && value <= 'Z'
               || value >= 'a' && value <= 'z';
    }

    private static int GetOrderedListSeparatorIndex(string line, int markerIndex) {
        int index = markerIndex;
        int digitCount = 0;
        while (index < line.Length && char.IsDigit(line[index]) && digitCount < 9) {
            index++;
            digitCount++;
        }

        if (digitCount == 0 || index >= line.Length || (line[index] != '.' && line[index] != ')')) {
            return -1;
        }

        int afterSeparator = index + 1;
        return afterSeparator >= line.Length || char.IsWhiteSpace(line[afterSeparator])
            ? index
            : -1;
    }

    private static int GetDefinitionListSeparatorIndex(string line, int markerIndex) {
        for (int index = markerIndex + 1; index < line.Length; index++) {
            if (line[index] != ':') {
                continue;
            }

            int afterSeparator = index + 1;
            return afterSeparator >= line.Length || char.IsWhiteSpace(line[afterSeparator])
                ? index
                : -1;
        }

        return -1;
    }

    private static bool StartsWithReferenceDefinitionLikeLabel(string line, int markerIndex, int colonIndex) {
        if (markerIndex < 0 || colonIndex <= markerIndex || colonIndex > line.Length) {
            return false;
        }

        string label = line.Substring(markerIndex, colonIndex - markerIndex).TrimEnd();
        if (label.Length == 0) {
            return false;
        }

        label = label.Replace(@"\*", "*")
            .Replace(@"\[", "[")
            .Replace(@"\]", "]");

        return (label.Length >= 2 && label[0] == '[' && label[label.Length - 1] == ']')
               || (label.Length >= 3 && label[0] == '*' && label[1] == '[' && label[label.Length - 1] == ']');
    }

    private static string EscapeTitleContent(string text, char delimiter) {
        if (string.IsNullOrEmpty(text)) return string.Empty;

        // Protect delimiter. (The title content is already escaped for common Markdown punctuation.)
        StringBuilder? sb = null;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (c == delimiter) {
                sb ??= new StringBuilder(text.Length + 4);
                sb.Append('\\').Append(c);
            } else {
                sb?.Append(c);
            }
        }
        return sb?.ToString() ?? text;
    }

    private static string EncodeLiteralMarkdownText(string? text, bool encodeEntityLikeAmpersands = true) {
        string value = text ?? string.Empty;
        if (value.Length == 0) {
            return string.Empty;
        }

        StringBuilder? sb = null;
        for (int i = 0; i < value.Length; i++) {
            string? replacement = value[i] switch {
                '\\' => @"\\",
                '[' => @"\[",
                ']' => @"\]",
                '(' => @"\(",
                ')' => @"\)",
                '|' => @"\|",
                '*' => @"\*",
                '_' => @"\_",
                '`' => @"\`",
                '~' => @"\~",
                '=' => @"\=",
                '<' => "&lt;",
                '>' => "&gt;",
                '&' when encodeEntityLikeAmpersands && IsEntityLikeAmpersand(value, i) => "&amp;",
                _ => null
            };

            if (replacement == null) {
                sb?.Append(value[i]);
                continue;
            }

            sb ??= new StringBuilder(value.Length + 16).Append(value, 0, i);
            sb.Append(replacement);
        }

        return sb?.ToString() ?? value;
    }

    private static bool IsEntityLikeAmpersand(string value, int index) {
        int semicolonIndex = value.IndexOf(';', index + 1);
        if (semicolonIndex < 0 || semicolonIndex - index > 32) {
            return false;
        }

        int entityStart = index + 1;
        if (entityStart >= semicolonIndex) {
            return false;
        }

        if (value[entityStart] == '#') {
            int numericStart = entityStart + 1;
            if (numericStart >= semicolonIndex) {
                return false;
            }

            if (value[numericStart] == 'x' || value[numericStart] == 'X') {
                int hexStart = numericStart + 1;
                if (hexStart >= semicolonIndex) {
                    return false;
                }

                for (int i = hexStart; i < semicolonIndex; i++) {
                    char c = value[i];
                    if (!((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))) {
                        return false;
                    }
                }

                return true;
            }

            for (int i = numericStart; i < semicolonIndex; i++) {
                if (!char.IsDigit(value[i])) {
                    return false;
                }
            }

            return true;
        }

        for (int i = entityStart; i < semicolonIndex; i++) {
            char c = value[i];
            if (!((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (i > entityStart && char.IsDigit(c)))) {
                return false;
            }
        }

        return true;
    }
}
