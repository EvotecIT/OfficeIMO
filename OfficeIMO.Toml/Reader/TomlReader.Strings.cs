namespace OfficeIMO.Toml;

public sealed partial class TomlReader {
    /// <summary>Scans a single-line basic string starting at its opening quote.</summary>
    private void ScanBasic(ReadOnlySpan<char> s, int start) {
        int i = start + 1;
        while (true) {
            if (i >= s.Length) throw Error(TomlErrorCodes.UnterminatedString, "The string is not closed.", start);
            char ch = s[i];
            if (ch == '"') break;
            if (ch == '\\') {
                int length = TomlScalars.EscapeLength(s, i);
                if (length == 0) throw Error(TomlErrorCodes.InvalidEscape, "The escape sequence is not valid.", i);
                i += length;
                continue;
            }
            if (ch == '\n' || ch == '\r') throw Error(TomlErrorCodes.UnterminatedString, "The string is not closed before the end of the line.", start);
            if (TomlScalars.IsForbiddenControl(ch)) throw Error(TomlErrorCodes.InvalidCharacter, "Control characters must be escaped in strings.", i);
            i++;
        }
        Finish(TomlTokenType.String, TomlStringStyle.Basic, start, i + 1);
    }

    /// <summary>Scans a single-line literal string starting at its opening quote.</summary>
    private void ScanLiteral(ReadOnlySpan<char> s, int start) {
        int i = start + 1;
        while (true) {
            if (i >= s.Length) throw Error(TomlErrorCodes.UnterminatedString, "The string is not closed.", start);
            char ch = s[i];
            if (ch == '\'') break;
            if (ch == '\n' || ch == '\r') throw Error(TomlErrorCodes.UnterminatedString, "The string is not closed before the end of the line.", start);
            if (TomlScalars.IsForbiddenControl(ch)) throw Error(TomlErrorCodes.InvalidCharacter, "Control characters are not allowed in literal strings.", i);
            i++;
        }
        Finish(TomlTokenType.String, TomlStringStyle.Literal, start, i + 1);
    }

    /// <summary>Scans a multi-line basic string starting at its opening <c>"""</c>.</summary>
    private void ScanMultilineBasic(ReadOnlySpan<char> s, int start) {
        int line = _line, lineStart = _lineStart;
        int i = start + 3;
        while (true) {
            if (i >= s.Length) throw Error(TomlErrorCodes.UnterminatedString, "The multi-line string is not closed.", start);
            char ch = s[i];
            if (ch == '"') {
                int run = QuoteRun(s, i, '"');
                if (run >= 3) {
                    if (run > 5) throw Error(TomlErrorCodes.UnexpectedToken, "A multi-line string can end with at most two quotes before its closing delimiter.", i + 5);
                    i += run;
                    break;
                }
                i += run;
                continue;
            }
            if (ch == '\\') {
                if (i + 1 < s.Length && (s[i + 1] == ' ' || s[i + 1] == '\t' || s[i + 1] == '\n' || s[i + 1] == '\r')) {
                    // A line-ending backslash: only whitespace may follow it on its line.
                    int j = i + 1;
                    while (j < s.Length && (s[j] == ' ' || s[j] == '\t')) j++;
                    if (j >= s.Length || (s[j] != '\n' && !(s[j] == '\r' && j + 1 < s.Length && s[j + 1] == '\n'))) {
                        throw Error(TomlErrorCodes.InvalidEscape, "Only whitespace may follow a line-ending backslash.", i);
                    }
                    i = j;
                    continue;
                }
                int length = TomlScalars.EscapeLength(s, i);
                if (length == 0) throw Error(TomlErrorCodes.InvalidEscape, "The escape sequence is not valid.", i);
                i += length;
                continue;
            }
            i = ScanMultilineChar(s, i, ch);
        }
        FinishMultiline(TomlStringStyle.MultilineBasic, start, i, line, lineStart);
    }

    /// <summary>Scans a multi-line literal string starting at its opening <c>'''</c>.</summary>
    private void ScanMultilineLiteral(ReadOnlySpan<char> s, int start) {
        int line = _line, lineStart = _lineStart;
        int i = start + 3;
        while (true) {
            if (i >= s.Length) throw Error(TomlErrorCodes.UnterminatedString, "The multi-line string is not closed.", start);
            char ch = s[i];
            if (ch == '\'') {
                int run = QuoteRun(s, i, '\'');
                if (run >= 3) {
                    if (run > 5) throw Error(TomlErrorCodes.UnexpectedToken, "A multi-line string can end with at most two quotes before its closing delimiter.", i + 5);
                    i += run;
                    break;
                }
                i += run;
                continue;
            }
            i = ScanMultilineChar(s, i, ch);
        }
        FinishMultiline(TomlStringStyle.MultilineLiteral, start, i, line, lineStart);
    }

    /// <summary>Checks one content character of a multi-line string, counting lines, and returns the next index.</summary>
    private int ScanMultilineChar(ReadOnlySpan<char> s, int i, char ch) {
        if (ch == '\n') {
            _line++;
            _lineStart = i + 1;
            return i + 1;
        }
        if (ch == '\r') {
            if (i + 1 >= s.Length || s[i + 1] != '\n') throw Error(TomlErrorCodes.InvalidCharacter, "A carriage return must be followed by a line feed.", i);
            _line++;
            _lineStart = i + 2;
            return i + 2;
        }
        if (TomlScalars.IsForbiddenControl(ch)) throw Error(TomlErrorCodes.InvalidCharacter, "Control characters are not allowed in strings.", i);
        return i + 1;
    }

    private static int QuoteRun(ReadOnlySpan<char> s, int i, char quote) {
        int run = 0;
        while (i + run < s.Length && s[i + run] == quote) run++;
        return run;
    }

    private void Finish(TomlTokenType type, TomlStringStyle style, int start, int end) {
        if (end - start > _options.MaxStringLength) throw new TomlLimitExceededException(nameof(TomlReaderOptions.MaxStringLength), _options.MaxStringLength, Position(start));
        _style = style;
        Token(type, start, end - start);
    }

    private void FinishMultiline(TomlStringStyle style, int start, int end, int line, int lineStart) {
        if (end - start > _options.MaxStringLength) throw new TomlLimitExceededException(nameof(TomlReaderOptions.MaxStringLength), _options.MaxStringLength, new TomlPosition(start, line, start - lineStart + 1));
        _style = style;
        _pos = end;
        Token(TomlTokenType.String, start, line, lineStart);
    }
}
