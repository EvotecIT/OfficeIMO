namespace OfficeIMO.Markdown;

using System;
using System.Text;

/// <summary>
/// Shared helper for escaping Markdown-reserved characters across inline renderers.
/// </summary>
internal static class MarkdownEscaper {
    // Inline HTML is allowed in rendered Markdown (e.g., <u> for underline), so we intentionally
    // do not escape angle brackets here to preserve legitimate HTML passthroughs.
    // Allows: "Use <u>underline</u> for emphasis" -> "Use <u>underline</u> for emphasis"
    // Escapes: "Text [link](url)" -> "Text \\[link\\]\(url\)"
    private static readonly char[] GeneralReserved = ['\\', '[', ']', '(', ')', '|', '*', '_'];
    private static readonly char[] UrlReserved = ['\\', '(', ')', '[', ']', '|'];

    internal static string EscapeText(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeEmphasis(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeLinkText(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeLinkUrl(string? text) => Escape(text, UrlReserved);
    internal static string EscapeImageAlt(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeImageSrc(string? text) => Escape(text, UrlReserved);

    internal static string FormatOptionalTitle(string? title) {
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
}
