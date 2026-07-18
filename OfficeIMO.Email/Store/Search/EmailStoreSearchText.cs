using OfficeIMO.Rtf;

namespace OfficeIMO.Email.Store;

internal static class EmailStoreSearchText {
    internal static string Normalize(string? value, int maximumCharacters) {
        if (string.IsNullOrEmpty(value) || maximumCharacters <= 0) return string.Empty;
        var builder = new StringBuilder(Math.Min(value!.Length, maximumCharacters));
        bool pendingSpace = false;
        for (int index = 0; index < value.Length && builder.Length < maximumCharacters; index++) {
            char character = value[index];
            if (char.IsWhiteSpace(character) || char.IsControl(character)) {
                pendingSpace = builder.Length > 0;
                continue;
            }
            if (pendingSpace && builder.Length < maximumCharacters) builder.Append(' ');
            pendingSpace = false;
            if (builder.Length < maximumCharacters) builder.Append(character);
        }
        return builder.ToString();
    }

    internal static string HtmlToText(string? html, int maximumCharacters) {
        if (string.IsNullOrEmpty(html) || maximumCharacters <= 0) return string.Empty;
        var visible = new StringBuilder(Math.Min(html!.Length, maximumCharacters));
        int index = 0;
        while (index < html.Length && visible.Length < maximumCharacters) {
            if (html[index] == '<') {
                int close = html.IndexOf('>', index + 1);
                if (close < 0) break;
                string tag = ReadTagName(html, index + 1, close);
                bool closingTag = IsClosingTag(html, index + 1, close);
                if (!closingTag &&
                    (string.Equals(tag, "script", StringComparison.OrdinalIgnoreCase) ||
                     string.Equals(tag, "style", StringComparison.OrdinalIgnoreCase))) {
                    int end = html.IndexOf(string.Concat("</", tag), close + 1,
                        StringComparison.OrdinalIgnoreCase);
                    if (end < 0) break;
                    index = end;
                    continue;
                }
                AppendSpace(visible, maximumCharacters);
                index = close + 1;
                continue;
            }
            if (html[index] == '&' && TryDecodeEntity(html, index, out char decoded, out int consumed)) {
                visible.Append(decoded);
                index += consumed;
                continue;
            }
            visible.Append(html[index++]);
        }
        return Normalize(visible.ToString(), maximumCharacters);
    }

    internal static string CreateSnippet(string text, IReadOnlyList<string> terms, int maximumCharacters) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        int matchIndex = int.MaxValue;
        int matchLength = 0;
        foreach (string term in terms) {
            int found = text.IndexOf(term, StringComparison.OrdinalIgnoreCase);
            if (found >= 0 && found < matchIndex) {
                matchIndex = found;
                matchLength = term.Length;
            }
        }
        if (matchIndex == int.MaxValue) return text.Length <= maximumCharacters
            ? text
            : text.Substring(0, maximumCharacters) + "…";

        int context = Math.Max(0, (maximumCharacters - matchLength) / 2);
        int start = Math.Max(0, matchIndex - context);
        int length = Math.Min(maximumCharacters, text.Length - start);
        string snippet = text.Substring(start, length);
        if (start > 0) snippet = "…" + snippet;
        if (start + length < text.Length) snippet += "…";
        return snippet;
    }

    internal static string RtfToText(string? rtf, int maximumCharacters,
        CancellationToken cancellationToken) {
        if (string.IsNullOrEmpty(rtf) || maximumCharacters <= 0) return string.Empty;
        try {
            RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();
            options.MaxInputCharacters = rtf!.Length;
            RtfDocument document = RtfDocument.Read(rtf, options, cancellationToken).Document;
            var builder = new StringBuilder(Math.Min(rtf.Length, maximumCharacters));
            AppendRtfBlocks(document.Blocks, builder, maximumCharacters);
            return Normalize(builder.ToString(), maximumCharacters);
        } catch (OperationCanceledException) {
            throw;
        } catch {
            // Searching preserved source is preferable to dropping the field when an uncommon RTF cannot be parsed.
            return Normalize(rtf, maximumCharacters);
        }
    }

    private static void AppendRtfBlocks(IReadOnlyList<IRtfBlock> blocks,
        StringBuilder builder, int maximumCharacters) {
        for (int index = 0; index < blocks.Count && builder.Length < maximumCharacters; index++) {
            switch (blocks[index]) {
                case RtfParagraph paragraph:
                    AppendRtfText(paragraph.ToPlainText(), builder, maximumCharacters);
                    break;
                case RtfTable table:
                    foreach (RtfTableRow row in table.Rows) {
                        foreach (RtfTableCell cell in row.Cells) {
                            AppendRtfBlocks(cell.Blocks, builder, maximumCharacters);
                        }
                    }
                    break;
                case RtfObject rtfObject:
                    AppendRtfText(rtfObject.ToPlainText(), builder, maximumCharacters);
                    break;
                case RtfShape shape:
                    AppendRtfText(shape.ToPlainText(), builder, maximumCharacters);
                    break;
            }
        }
    }

    private static void AppendRtfText(string? value, StringBuilder builder, int maximumCharacters) {
        if (string.IsNullOrWhiteSpace(value) || builder.Length >= maximumCharacters) return;
        AppendSpace(builder, maximumCharacters);
        int available = maximumCharacters - builder.Length;
        builder.Append(value!.Length <= available ? value : value.Substring(0, available));
    }

    private static string ReadTagName(string html, int start, int close) {
        while (start < close && (char.IsWhiteSpace(html[start]) || html[start] == '/')) start++;
        int end = start;
        while (end < close && (char.IsLetterOrDigit(html[end]) || html[end] == ':' || html[end] == '-')) end++;
        return end > start ? html.Substring(start, end - start) : string.Empty;
    }

    private static bool IsClosingTag(string html, int start, int close) {
        while (start < close && char.IsWhiteSpace(html[start])) start++;
        return start < close && html[start] == '/';
    }

    private static bool TryDecodeEntity(string value, int start, out char decoded, out int consumed) {
        decoded = default;
        consumed = 0;
        int semicolon = value.IndexOf(';', start + 1);
        if (semicolon < 0 || semicolon - start > 10) return false;
        string entity = value.Substring(start + 1, semicolon - start - 1);
        if (string.Equals(entity, "amp", StringComparison.OrdinalIgnoreCase)) decoded = '&';
        else if (string.Equals(entity, "lt", StringComparison.OrdinalIgnoreCase)) decoded = '<';
        else if (string.Equals(entity, "gt", StringComparison.OrdinalIgnoreCase)) decoded = '>';
        else if (string.Equals(entity, "quot", StringComparison.OrdinalIgnoreCase)) decoded = '"';
        else if (string.Equals(entity, "apos", StringComparison.OrdinalIgnoreCase)) decoded = '\'';
        else if (string.Equals(entity, "nbsp", StringComparison.OrdinalIgnoreCase)) decoded = ' ';
        else if (entity.Length > 1 && entity[0] == '#') {
            bool hexadecimal = entity.Length > 2 && (entity[1] == 'x' || entity[1] == 'X');
            string number = entity.Substring(hexadecimal ? 2 : 1);
            if (!int.TryParse(number,
                hexadecimal ? NumberStyles.AllowHexSpecifier : NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int scalar) || scalar < 0 || scalar > char.MaxValue) return false;
            decoded = (char)scalar;
        } else {
            return false;
        }
        consumed = semicolon - start + 1;
        return true;
    }

    private static void AppendSpace(StringBuilder builder, int maximumCharacters) {
        if (builder.Length > 0 && builder.Length < maximumCharacters && builder[builder.Length - 1] != ' ') {
            builder.Append(' ');
        }
    }
}
