using System;
using System.Text;

namespace OfficeIMO.Rtf.Markdown;

internal static class RtfMarkdownText {
    internal static string EscapeMarkdownText(string? text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        var sb = new StringBuilder(text!.Length + 8);
        for (int i = 0; i < text.Length; i++) {
            char ch = text[i];
            switch (ch) {
                case '\\':
                case '`':
                case '*':
                case '_':
                case '{':
                case '}':
                case '[':
                case ']':
                case '(':
                case ')':
                case '#':
                case '+':
                case '-':
                case '.':
                case '!':
                case '|':
                case '~':
                    sb.Append('\\');
                    break;
            }

            sb.Append(ch);
        }

        return sb.ToString();
    }

    internal static string EscapeImageAlt(string? text) {
        return EscapeMarkdownText(text).Replace("\r", " ").Replace("\n", " ");
    }

    internal static string EscapeLinkUrl(string? url) {
        if (string.IsNullOrEmpty(url)) {
            return string.Empty;
        }

        return url!
            .Replace("\\", "%5C")
            .Replace(" ", "%20")
            .Replace(")", "%29");
    }

    internal static string HtmlEncode(string? text) {
        return System.Net.WebUtility.HtmlEncode(text ?? string.Empty);
    }

    internal static string PlainText(OfficeIMO.Markdown.IMarkdownInline inline) {
        if (inline is OfficeIMO.Markdown.IPlainTextMarkdownInline plainText) {
            var sb = new StringBuilder();
            plainText.AppendPlainText(sb);
            return sb.ToString();
        }

        if (inline is OfficeIMO.Markdown.IRenderableMarkdownInline renderable) {
            return renderable.RenderMarkdown();
        }

        return inline.ToString() ?? string.Empty;
    }
}
