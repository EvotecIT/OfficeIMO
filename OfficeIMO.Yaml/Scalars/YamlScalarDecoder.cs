namespace OfficeIMO.Yaml;

/// <summary>
/// Turns a scalar's raw source text into its value: line folding, escapes, block indentation and chomping.
/// The scanner has already validated the text, so decoding cannot fail. Decoded text is never longer than the raw text.
/// </summary>
internal static class YamlScalarDecoder {
    /// <summary>Returns whether the raw content span is already the scalar's value.</summary>
    public static bool IsVerbatim(YamlStyle style, YamlTokenFlags flags) =>
        (style == YamlStyle.Plain || style == YamlStyle.SingleQuoted || style == YamlStyle.DoubleQuoted)
        && (flags & (YamlTokenFlags.MultiLine | YamlTokenFlags.Escapes)) == 0;

    /// <summary>Decodes <paramref name="raw"/> into <paramref name="destination"/>, which must be at least as long, and returns the length.</summary>
    public static int Decode(ReadOnlySpan<char> raw, YamlStyle style, YamlChomping chomping, int indent, Span<char> destination) {
        switch (style) {
            case YamlStyle.Literal:
            case YamlStyle.Folded:
                return DecodeBlock(raw, style == YamlStyle.Folded, chomping, indent, destination);
            case YamlStyle.DoubleQuoted:
                return DecodeFlow(raw, doubleQuoted: true, singleQuoted: false, destination);
            case YamlStyle.SingleQuoted:
                return DecodeFlow(raw, doubleQuoted: false, singleQuoted: true, destination);
            default:
                // Plain scalars have no escapes: '' stays two characters.
                return DecodeFlow(raw, doubleQuoted: false, singleQuoted: false, destination);
        }
    }

    /// <summary>Folds plain and quoted scalars and expands escapes (<c>''</c> in single quotes, backslash escapes in double quotes).</summary>
    private static int DecodeFlow(ReadOnlySpan<char> raw, bool doubleQuoted, bool singleQuoted, Span<char> o) {
        int n = raw.Length;
        int w = 0;
        int keep = 0;
        int i = 0;
        while (i < n) {
            char ch = raw[i];
            if (ch == '\'' && singleQuoted && i + 1 < n && raw[i + 1] == '\'') {
                o[w++] = '\'';
                keep = w;
                i += 2;
                continue;
            }
            if (ch == '\\' && doubleQuoted) {
                char e = raw[i + 1];
                if (e == '\r' || e == '\n') {
                    i += 1;
                    i += e == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
                    keep = w;
                    // Escaped line break: drop it and the following line prefix; empty lines still count.
                    while (true) {
                        while (i < n && (raw[i] == ' ' || raw[i] == '\t')) i++;
                        if (i < n && (raw[i] == '\r' || raw[i] == '\n')) {
                            i += raw[i] == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
                            o[w++] = '\n';
                            keep = w;
                            continue;
                        }
                        break;
                    }
                    continue;
                }
                i += 2;
                int code;
                switch (e) {
                    case '0': code = 0; break;
                    case 'a': code = 7; break;
                    case 'b': code = 8; break;
                    case 't': case '\t': code = 9; break;
                    case 'n': code = 10; break;
                    case 'v': code = 11; break;
                    case 'f': code = 12; break;
                    case 'r': code = 13; break;
                    case 'e': code = 27; break;
                    case 'N': code = 0x85; break;
                    case '_': code = 0xA0; break;
                    case 'L': code = 0x2028; break;
                    case 'P': code = 0x2029; break;
                    case 'x': code = ParseHex(raw, i, 2); i += 2; break;
                    case 'u': code = ParseHex(raw, i, 4); i += 4; break;
                    case 'U': code = ParseHex(raw, i, 8); i += 8; break;
                    default: code = e; break;
                }
                if (code >= 0x10000) {
                    code -= 0x10000;
                    o[w++] = (char)(0xD800 + (code >> 10));
                    o[w++] = (char)(0xDC00 + (code & 0x3FF));
                } else {
                    o[w++] = (char)code;
                }
                keep = w;
                continue;
            }
            if (ch == '\r' || ch == '\n') {
                w = keep;
                i += ch == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
                int empty = 0;
                while (true) {
                    while (i < n && (raw[i] == ' ' || raw[i] == '\t')) i++;
                    if (i < n && (raw[i] == '\r' || raw[i] == '\n')) {
                        i += raw[i] == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
                        empty++;
                        continue;
                    }
                    break;
                }
                if (empty == 0) o[w++] = ' ';
                else for (int k = 0; k < empty; k++) o[w++] = '\n';
                keep = w;
                continue;
            }
            o[w++] = ch;
            if (ch != ' ' && ch != '\t') keep = w;
            i++;
        }
        return w;
    }

    private static int ParseHex(ReadOnlySpan<char> raw, int start, int count) {
        int value = 0;
        for (int k = 0; k < count; k++) value = value * 16 + YamlChars.HexValue(raw[start + k]);
        return value;
    }

    /// <summary>
    /// Decodes block scalar content. <paramref name="raw"/> starts at the first line after the header; with clip and
    /// strip chomping it ends after the last content line, with keep chomping it includes trailing empty lines.
    /// </summary>
    private static int DecodeBlock(ReadOnlySpan<char> raw, bool folded, YamlChomping chomping, int indent, Span<char> o) {
        int n = raw.Length;
        int w = 0;
        int i = 0;
        bool leadingBreak = false;     // a line break follows the previous content line
        int trailingBreaks = 0;        // empty lines since the previous content line
        bool leadingBlank = false;     // the previous content line starts with whitespace
        bool first = true;
        while (i < n) {
            // Skip up to the indentation.
            int spaces = 0;
            while (spaces < indent && i < n && raw[i] == ' ') { i++; spaces++; }
            if (i >= n) {
                // A final line of spaces without a line break is an empty line.
                if (spaces > 0) trailingBreaks++;
                break;
            }
            char ch = raw[i];
            if (ch == '\r' || ch == '\n') {
                i += ch == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
                trailingBreaks++;
                continue;
            }
            // A content line.
            bool trailingBlank = ch == ' ' || ch == '\t';
            if (!first && folded && leadingBreak && !leadingBlank && !trailingBlank) {
                if (trailingBreaks == 0) o[w++] = ' ';
            } else if (leadingBreak) {
                o[w++] = '\n';
            }
            for (int k = 0; k < trailingBreaks; k++) o[w++] = '\n';
            trailingBreaks = 0;
            leadingBreak = false;
            leadingBlank = trailingBlank;
            first = false;
            bool blankLine = true;
            while (i < n && raw[i] != '\r' && raw[i] != '\n') {
                if (raw[i] != ' ' && raw[i] != '\t') blankLine = false;
                o[w++] = raw[i++];
            }
            // A last content line that ends at the end of input has no line break to keep (YAML 1.2.2 section 8.1.1.2);
            // a last line of only white space still counts as ended, as the reference parsers read it (suite L24T).
            leadingBreak = i < n || blankLine;
            if (i < n) i += raw[i] == '\r' && i + 1 < n && raw[i + 1] == '\n' ? 2 : 1;
        }
        if (chomping != YamlChomping.Strip && leadingBreak) o[w++] = '\n';
        if (chomping == YamlChomping.Keep) {
            for (int k = 0; k < trailingBreaks; k++) o[w++] = '\n';
        }
        return w;
    }
}
