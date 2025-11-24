namespace OfficeIMO.Markdown;

using System;
using System.Text;

/// <summary>
/// Shared helper for escaping Markdown-reserved characters across inline renderers.
/// </summary>
internal static class MarkdownEscaper {
    // Inline HTML is allowed in rendered Markdown (e.g., <u> for underline), so we intentionally
    // do not escape angle brackets here to preserve legitimate HTML passthroughs.
    private static readonly char[] GeneralReserved = ['\\', '[', ']', '(', ')', '|', '*', '_'];
    private static readonly char[] UrlReserved = ['\\', '(', ')', '[', ']', '|'];

    internal static string EscapeText(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeEmphasis(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeLinkText(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeLinkUrl(string? text) => Escape(text, UrlReserved);
    internal static string EscapeImageAlt(string? text) => Escape(text, GeneralReserved);
    internal static string EscapeImageSrc(string? text) => Escape(text, UrlReserved);

    private static string Escape(string? text, char[] reserved) {
        if (string.IsNullOrEmpty(text)) return string.Empty;

        StringBuilder sb = new(text.Length);
        foreach (char c in text) {
            if (Array.IndexOf(reserved, c) >= 0) sb.Append('\\');
            sb.Append(c);
        }
        return sb.ToString();
    }
}
