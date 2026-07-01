using System.Text.RegularExpressions;
using OfficeIMO.Markdown;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class CommonMarkHtmlComparison {
    public static HtmlOptions CreatePlainHtmlOptions() {
        return new HtmlOptions {
            Style = HtmlStyle.Plain,
            CssDelivery = CssDelivery.None,
            BodyClass = null
        };
    }

    public static string Normalize(string html) {
        if (string.IsNullOrWhiteSpace(html)) {
            return string.Empty;
        }

        var sb = new StringBuilder(html.Length);
        bool inTag = false;
        bool lastWasWhitespace = false;

        for (int i = 0; i < html.Length; i++) {
            char ch = html[i];
            if (ch == '<') {
                if (!inTag && lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                    sb.Append(' ');
                }

                inTag = true;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (ch == '>') {
                inTag = false;
                lastWasWhitespace = false;
                sb.Append(ch);
                continue;
            }

            if (inTag) {
                sb.Append(ch);
                continue;
            }

            if (IsHtmlSpace(ch)) {
                lastWasWhitespace = true;
                continue;
            }

            if (lastWasWhitespace && sb.Length > 0 && sb[sb.Length - 1] != '>') {
                sb.Append(' ');
            }

            lastWasWhitespace = false;
            sb.Append(ch);
        }

        string normalized = sb.ToString()
            .Replace("> <", "><")
            .Replace("<br/>", "<br />")
            .Replace("&#39;", "'")
            .Replace("&#x27;", "'");
        normalized = Regex.Replace(
            normalized,
            "&#(\\d+);",
            static match => char.ConvertFromUtf32(int.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture)),
            RegexOptions.CultureInvariant);
        normalized = Regex.Replace(
            normalized,
            "&#x([0-9a-fA-F]+);",
            static match => char.ConvertFromUtf32(int.Parse(match.Groups[1].Value, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture)),
            RegexOptions.CultureInvariant);
        normalized = Regex.Replace(normalized, "<h([1-6])\\s+id=\"[^\"]*\">", "<h$1>", RegexOptions.CultureInvariant);
        normalized = normalized
            .Replace(" <ul", "<ul")
            .Replace(" <ol", "<ol")
            .Replace(" <blockquote", "<blockquote")
            .Replace(" <pre", "<pre")
            .Replace(" <table", "<table")
            .Replace(" <p", "<p");

        return normalized.Trim();
    }

    private static bool IsHtmlSpace(char ch) =>
        ch == ' ' || ch == '\t' || ch == '\n' || ch == '\r' || ch == '\f';
}
