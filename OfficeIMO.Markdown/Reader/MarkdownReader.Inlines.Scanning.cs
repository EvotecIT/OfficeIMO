namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
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
        return lookup;
    }

    private static bool IsBackslashEscapable(char c) {
        // CommonMark backslash-escapable punctuation (plus '|' which we want for tables).
        // See: https://spec.commonmark.org/ (backslash escapes). We keep the set small and pragmatic.
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
            '+' => true,
            '-' => true,
            '.' => true,
            '!' => true,
            '"' => true,
            '\'' => true,
            '|' => true,
            '>' => true,
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
}
