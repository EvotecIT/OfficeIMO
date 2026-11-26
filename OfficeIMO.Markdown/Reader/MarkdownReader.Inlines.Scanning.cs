namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    private static readonly bool[] PotentialInlineStartLookup = CreatePotentialInlineStartLookup();

    private static bool[] CreatePotentialInlineStartLookup() {
        var lookup = new bool[128];
        lookup['['] = true;
        lookup['!'] = true;
        lookup['`'] = true;
        lookup['*'] = true;
        lookup['_'] = true;
        lookup['~'] = true;
        lookup['='] = true;
        lookup['+'] = true;
        lookup['^'] = true;
        return lookup;
    }

    // Characters at which a plain text run may end: the potential inline starts, line feeds,
    // backslash escapes, entities and angle brackets.
    private static readonly char[] PlainTextRunStops = { '[', '!', '`', '*', '_', '~', '=', '+', '^', '\n', '\\', '&', '<' };

    /// <summary>
    /// Finds where a text run that started before <paramref name="position"/> ends when no autolink,
    /// abbreviation or parser extension can start inside it: at the first potential inline start, line feed,
    /// escaped character, entity, or angle bracket that starts inline HTML or an angle autolink.
    /// </summary>
    private static int FindPlainTextRunEnd(string text, int position, bool allowInlineHtml, bool allowLinks, bool allowImages) {
        while (position < text.Length) {
            int next = text.IndexOfAny(PlainTextRunStops, position);
            if (next < 0) {
                return text.Length;
            }

            position = next;
            char value = text[position];
            switch (value) {
                case '\\':
                    if (position + 1 < text.Length && IsBackslashEscapable(text[position + 1])) {
                        return position;
                    }
                    break;
                case '&':
                    if (TryConsumeHtmlEntityText(text, position, out _, out _)) {
                        return position;
                    }
                    break;
                case '<':
                    if (allowInlineHtml || IsAngleAutolinkStart(text, position)) {
                        return position;
                    }
                    break;
                case '[' when !allowLinks:
                case '!' when !allowImages:
                    break;
                default:
                    return position;
            }

            position++;
        }

        return position;
    }

    private static bool IsBackslashEscapable(char c) {
        // CommonMark backslash-escapable ASCII punctuation.
        return c switch {
            '\\' => true,
            '`' => true,
            '*' => true,
            '_' => true,
            '{' => true,
            '}' => true,
            '[' => true,
            ']' => true,
            '(' => true,
            ')' => true,
            '#' => true,
            '$' => true,
            '%' => true,
            '&' => true,
            '+' => true,
            ',' => true,
            '-' => true,
            '.' => true,
            '/' => true,
            ':' => true,
            ';' => true,
            '<' => true,
            '!' => true,
            '"' => true,
            '\'' => true,
            '|' => true,
            '>' => true,
            '?' => true,
            '@' => true,
            '^' => true,
            '~' => true,
            '=' => true,
            _ => false
        };
    }

    private static bool IsIntrawordDelimiter(string text, int start, int markerLength) {
        // Pragmatic GFM-ish rule: treat '_' emphasis markers as disabled when they appear inside "words".
        // This avoids accidentally italicizing identifiers like foo_bar_baz.
        if (string.IsNullOrEmpty(text)) return false;
        int left = start - 1;
        int right = start + markerLength;
        if (left < 0 || right >= text.Length) return false;
        return char.IsLetterOrDigit(text[left]) && char.IsLetterOrDigit(text[right]);
    }

    private static bool IsPotentialInlineStart(char c, bool allowInlineHtml, bool allowLinks, bool allowImages) {
        if (allowInlineHtml && c == '<') return true;
        if (c < PotentialInlineStartLookup.Length && PotentialInlineStartLookup[c]) {
            if (!allowLinks && c == '[') return false;
            if (!allowImages && c == '!') return false;
            return true;
        }
        return false;
    }

    private static bool ContainsPotentialInlineSyntax(
        string text,
        MarkdownLoadOptions options,
        bool allowLinks,
        bool allowImages) {
        for (int i = 0; i < text.Length; i++) {
            char value = text[i];
            if (value == '\\' || value == '&' || value == '\n' ||
                IsPotentialInlineStart(value, options.InlineHtml, allowLinks, allowImages)) {
                return true;
            }

            if (value == '$' && options.Math) {
                return true;
            }
        }

        return allowLinks && ContainsPotentialBareAutolinkSyntax(text, options);
    }

    private static bool ContainsPotentialBareAutolinkSyntax(string text, MarkdownLoadOptions options) {
        if (string.IsNullOrEmpty(text)) {
            return false;
        }

        if (options.AutolinkUrls
            && (text.IndexOf("http://", StringComparison.Ordinal) >= 0
                || text.IndexOf("https://", StringComparison.Ordinal) >= 0)) {
            return true;
        }

        if (options.AutolinkWwwUrls) {
            var comparison = options.AutolinkRequireLowercaseWwwPrefix
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;
            if (text.IndexOf("www.", comparison) >= 0) {
                return true;
            }
        }

        if (options.AutolinkBareSchemeUrls
            && (ContainsEnabledBareScheme(text, options, "mailto:")
                || ContainsEnabledBareScheme(text, options, "ftp://")
                || ContainsEnabledBareScheme(text, options, "tel:")
                || ContainsEnabledBareScheme(text, options, "xmpp:"))) {
            return true;
        }

        // A plain email autolink cannot exist without an at sign. The full parser
        // performs the precise boundary and address validation when one is present.
        return options.AutolinkEmails && text.IndexOf('@') >= 0;
    }

    private static bool ContainsEnabledBareScheme(
        string text,
        MarkdownLoadOptions options,
        string scheme) {
        if (!IsBareSchemePrefixEnabled(options, scheme)) {
            return false;
        }

        var comparison = options.AutolinkRequireLowercaseBareSchemePrefix
            ? StringComparison.Ordinal
            : StringComparison.OrdinalIgnoreCase;
        return text.IndexOf(scheme, comparison) >= 0;
    }
}
