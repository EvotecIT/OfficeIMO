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

#if NET8_0_OR_GREATER
    private static readonly System.Buffers.SearchValues<char> AsciiSpecialCharacterValues =
        System.Buffers.SearchValues.Create("&<>\"'");
    // ASCII characters that come out unchanged under either policy.
    private static readonly System.Buffers.SearchValues<char> UnchangedAsciiCharacterValues = CreateUnchangedAsciiCharacterValues();

    private static System.Buffers.SearchValues<char> CreateUnchangedAsciiCharacterValues() {
        var characters = new System.Collections.Generic.List<char>(128);
        for (char value = (char)0; value < (char)128; value++) {
            if (value is not ('&' or '<' or '>' or '"' or '\'')) {
                characters.Add(value);
            }
        }

        return System.Buffers.SearchValues.Create(characters.ToArray());
    }
#endif

    /// <summary>
    /// Appends the <paramref name="length"/> characters of <paramref name="text"/> at <paramref name="start"/>
    /// encoded as <see cref="Append(StringBuilder, string?, MarkdownHtmlExportOptions?)"/> encodes them as a string.
    /// </summary>
    internal static void Append(StringBuilder builder, string text, int start, int length, MarkdownHtmlExportOptions? options) {
        if (length <= 0) {
            return;
        }

        if (start == 0 && length == text.Length) {
            Append(builder, text, options);
            return;
        }

        int end = start + length;
        bool escapeNonAscii = options?.EscapeNonAsciiText != false;
        // The first character that is not plain ASCII text: one of the five special characters, or
        // with non-ASCII escaping any character from 128 up.
#if NET8_0_OR_GREATER
        int found = escapeNonAscii
            ? text.AsSpan(start, length).IndexOfAnyExcept(UnchangedAsciiCharacterValues)
            : text.AsSpan(start, length).IndexOfAny(AsciiSpecialCharacterValues);
        int first = found < 0 ? end : start + found;
#else
        int first = start;
        while (first < end) {
            char current = text[first];
            if (current is '&' or '<' or '>' or '"' or '\'' || (escapeNonAscii && current >= 128)) {
                break;
            }

            first++;
        }
#endif
        if (first == end) {
            builder.Append(text, start, length);
            return;
        }

        // The five special characters are written the same under both policies, so the range is escaped
        // in place unless a non-ASCII character needs WebUtility's numeric encoding.
        if (escapeNonAscii) {
            for (int i = first; i < end; i++) {
                if (text[i] >= 128) {
                    builder.Append(System.Net.WebUtility.HtmlEncode(text.Substring(start, length)));
                    return;
                }
            }
        }

        AppendEscaped(builder, text, start, end, first);
    }

    private static void AppendEscaped(StringBuilder builder, string value, int first) =>
        AppendEscaped(builder, value, 0, value.Length, first);

    private static void AppendEscaped(StringBuilder builder, string value, int start, int end, int first) {
        builder.Append(value, start, first - start);
        int runStart = first;
        for (var i = first; i < end; i++) {
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

        if (runStart < end) {
            builder.Append(value, runStart, end - runStart);
        }
    }
}
