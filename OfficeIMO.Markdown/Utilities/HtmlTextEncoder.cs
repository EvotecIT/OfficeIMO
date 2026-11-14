namespace OfficeIMO.Markdown;

internal static class HtmlTextEncoder {
    private static readonly char[] AsciiSpecialCharacters = { '&', '<', '>', '"', '\'' };

    internal static string Encode(string? text, MarkdownHtmlExportOptions? options) {
        if (options?.EscapeNonAsciiText == false) {
            return Encode(text);
        }

        // WebUtility returns the input itself when nothing needs encoding.
        return System.Net.WebUtility.HtmlEncode(text ?? string.Empty);
    }

    internal static string Encode(string? text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        var value = text!;
        int first = value.IndexOfAny(AsciiSpecialCharacters);
        if (first < 0) {
            return value;
        }

        var builder = new StringBuilder(value.Length + 16);
        AppendEscaped(builder, value, first);
        return builder.ToString();
    }

    /// <summary>Appends <paramref name="text"/> encoded with the same policy as <see cref="Encode(string?, MarkdownHtmlExportOptions?)"/>.</summary>
    internal static void Append(StringBuilder builder, string? text, MarkdownHtmlExportOptions? options) {
        if (string.IsNullOrEmpty(text)) {
            return;
        }

        if (options?.EscapeNonAsciiText == false) {
            int first = text!.IndexOfAny(AsciiSpecialCharacters);
            if (first < 0) {
                builder.Append(text);
            } else {
                AppendEscaped(builder, text, first);
            }

            return;
        }

        builder.Append(System.Net.WebUtility.HtmlEncode(text));
    }

    private static void AppendEscaped(StringBuilder builder, string value, int first) {
        builder.Append(value, 0, first);
        int runStart = first;
        for (var i = first; i < value.Length; i++) {
            string? replacement = value[i] switch {
                '&' => "&amp;",
                '<' => "&lt;",
                '>' => "&gt;",
                '"' => "&quot;",
                '\'' => "&#39;",
                _ => null
            };
            if (replacement == null) {
                continue;
            }

            if (i > runStart) {
                builder.Append(value, runStart, i - runStart);
            }

            builder.Append(replacement);
            runStart = i + 1;
        }

        if (runStart < value.Length) {
            builder.Append(value, runStart, value.Length - runStart);
        }
    }
}
