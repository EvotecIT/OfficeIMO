namespace OfficeIMO.Yaml;

internal ref partial struct YamlScanner {
    /// <summary>
    /// Scans a plain scalar. The token span excludes trailing whitespace; lookahead for continuation lines is
    /// rewound, so the gap after the token is rescanned as trivia.
    /// </summary>
    private YamlToken ScanPlainScalar() {
        YamlScanContext c = _c;
        bool flow = c.FlowLevel > 0;
        int indentLimit = c.Indent + 1;
        int start = c.Pos, startLine = c.Line, startColumn = c.Column;
        int pos = start, line = startLine, column = startColumn;
        int end = start, endLine = line, endColumn = column;
        int lineContentStart = start;
        bool multiLine = false;

        while (true) {
            bool stop = false;
            while (true) {
                int index = YamlChars.IndexOfPlainStop(_text.Slice(pos), flow);
                if (index < 0) {
                    int rest = _text.Length - pos;
                    pos += rest;
                    column += rest;
                    if (Fill()) continue;
                    break;
                }
                pos += index;
                column += index;
                char ch = _text[pos];
                if (ch == ':') {
                    char next = At(pos + 1);
                    if (YamlChars.IsBlankOrBreakOrNul(next) || (flow && YamlChars.IsFlowIndicator(next))) {
                        stop = true;
                        break;
                    }
                    pos++;
                    column++;
                    continue;
                }
                if (ch == '#') {
                    if (YamlChars.IsBlank(_text[pos - 1])) {
                        stop = true;
                        break;
                    }
                    pos++;
                    column++;
                    continue;
                }
                if (YamlChars.IsBreak(ch)) break;
                if (flow && YamlChars.IsFlowIndicator(ch)) {
                    stop = true;
                    break;
                }
                c.Pos = pos;
                c.Line = line;
                c.Column = column;
                throw Error(YamlErrorCodes.InvalidCharacter, "Found a character that is not allowed in a plain scalar: '" + Describe(ch) + "'.");
            }

            int e = pos;
            while (e > lineContentStart && YamlChars.IsBlank(_text[e - 1])) e--;
            if (e > lineContentStart) {
                end = e;
                endLine = line;
                endColumn = column - (pos - e);
            }
            if (stop || !HasChar(pos)) break;

            // At a line break: look for a continuation line.
            int p = pos;
            int l = line;
            int col;
            bool continues = false;
            while (true) {
                if (At(p) == '\r' && At(p + 1) == '\n') p += 2;
                else p++;
                l++;
                col = 0;
                while (At(p) == ' ') { p++; col++; }
                int spaces = col;
                while (YamlChars.IsBlank(At(p))) { p++; col++; }
                char ch = At(p);
                if (YamlChars.IsBreak(ch)) continue;
                if (ch == '\0' && !HasChar(p)) break;
                if (col == 0 && IsDocumentMarkerAt(p)) break;
                if (ch == '#') break;
                if (spaces < indentLimit) break;
                if (ch == ':') {
                    char next = At(p + 1);
                    if (YamlChars.IsBlankOrBreakOrNul(next) || (flow && YamlChars.IsFlowIndicator(next))) break;
                }
                if (flow && YamlChars.IsFlowIndicator(ch)) break;
                continues = true;
                break;
            }
            if (!continues) break;
            multiLine = true;
            pos = p;
            line = l;
            column = col;
            lineContentStart = p;
        }

        c.Pos = end;
        c.Line = endLine;
        c.Column = endColumn;
        CheckScalarLength(start, startLine, startColumn);
        YamlToken token = NewToken(YamlTokenKind.Scalar, start, startLine, startColumn);
        token.Style = YamlStyle.Plain;
        token.Flags = multiLine ? YamlTokenFlags.MultiLine : YamlTokenFlags.None;
        token.ValueStart = token.Start;
        token.ValueEnd = token.End;
        return token;
    }

    /// <summary>
    /// Scans a single- or double-quoted scalar and validates its escapes, so decoding later cannot fail.
    /// Continuation lines must be indented more than the enclosing block collection.
    /// </summary>
    private YamlToken ScanFlowScalar(bool single) {
        YamlScanContext c = _c;
        int start = c.Pos, startLine = c.Line, startColumn = c.Column;
        int minIndent = c.Indent + 1;
        YamlTokenFlags flags = YamlTokenFlags.None;
        Advance(1);
        int valueStart = c.Pos;
        while (true) {
            ReadOnlySpan<char> rest = _text.Slice(c.Pos);
            int index = single ? YamlChars.IndexOfSingleQuotedStop(rest) : YamlChars.IndexOfDoubleQuotedStop(rest);
            if (index < 0) {
                Advance(rest.Length);
                if (Fill()) continue;
                throw Error(YamlErrorCodes.UnterminatedScalar, "The quoted scalar is not closed.", Absolute(start), startLine, startColumn);
            }
            Advance(index);
            char ch = _text[c.Pos];
            if (single && ch == '\'') {
                if (At(c.Pos + 1) == '\'') {
                    flags |= YamlTokenFlags.Escapes;
                    Advance(2);
                    continue;
                }
                break;
            }
            if (!single && ch == '"') break;
            if (ch == '\\') {
                flags |= YamlTokenFlags.Escapes;
                ScanEscape();
                continue;
            }
            if (YamlChars.IsBreak(ch)) {
                flags |= YamlTokenFlags.MultiLine;
                AdvanceBreak();
                while (true) {
                    int spaces = 0;
                    while (At(c.Pos) == ' ') { Advance(1); spaces++; }
                    while (YamlChars.IsBlank(At(c.Pos))) Advance(1);
                    char next = At(c.Pos);
                    if (YamlChars.IsBreak(next)) {
                        AdvanceBreak();
                        continue;
                    }
                    if (next == '\0' && !HasChar(c.Pos)) {
                        throw Error(YamlErrorCodes.UnterminatedScalar, "The quoted scalar is not closed.", Absolute(start), startLine, startColumn);
                    }
                    if (c.Column == 0 && IsDocumentMarkerAt(c.Pos)) {
                        throw Error(YamlErrorCodes.InvalidDocument, "A document marker cannot appear inside a quoted scalar.");
                    }
                    if (spaces < minIndent) {
                        throw Error(YamlErrorCodes.InvalidIndentation, "A quoted scalar's continuation line must be indented more than the enclosing block collection.");
                    }
                    break;
                }
                continue;
            }
            throw Error(YamlErrorCodes.InvalidCharacter, "Found a character that is not allowed in a quoted scalar: '" + Describe(ch) + "'.");
        }
        int valueEnd = c.Pos;
        Advance(1);
        CheckScalarLength(start, startLine, startColumn);
        YamlToken token = NewToken(YamlTokenKind.Scalar, start, startLine, startColumn);
        token.Style = single ? YamlStyle.SingleQuoted : YamlStyle.DoubleQuoted;
        token.Flags = flags;
        token.ValueStart = Absolute(valueStart);
        token.ValueEnd = Absolute(valueEnd);
        return token;
    }

    private void ScanEscape() {
        YamlScanContext c = _c;
        char e = At(c.Pos + 1);
        int digits;
        switch (e) {
            case '0': case 'a': case 'b': case 't': case '\t': case 'n': case 'v': case 'f': case 'r': case 'e':
            case ' ': case '"': case '/': case '\\': case 'N': case '_': case 'L': case 'P':
                Advance(2);
                return;
            case '\r':
            case '\n':
                Advance(1);
                return;
            case 'x': digits = 2; break;
            case 'u': digits = 4; break;
            case 'U': digits = 8; break;
            default:
                throw Error(YamlErrorCodes.InvalidEscape, "Unknown escape sequence '\\" + Describe(e) + "' in a double-quoted scalar.");
        }
        long value = 0;
        for (int i = 0; i < digits; i++) {
            int hex = YamlChars.HexValue(At(c.Pos + 2 + i));
            if (hex < 0) throw Error(YamlErrorCodes.InvalidEscape, "The escape '\\" + e + "' needs " + digits.ToString(CultureInfo.InvariantCulture) + " hexadecimal digits.");
            value = value * 16 + hex;
        }
        if (value > 0x10FFFF) throw Error(YamlErrorCodes.InvalidEscape, "The escape '\\U' is beyond the Unicode range.");
        Advance(2 + digits);
    }

    /// <summary>
    /// Scans a literal or folded block scalar. The token's value span covers the content lines; with clip and strip
    /// chomping the token ends after the last content line, so trailing blank lines stay in the following gap.
    /// </summary>
    private YamlToken ScanBlockScalar(YamlStyle style) {
        YamlScanContext c = _c;
        int start = c.Pos, startLine = c.Line, startColumn = c.Column;
        Advance(1);
        YamlChomping chomping = YamlChomping.Clip;
        int increment = 0;
        char ch = At(c.Pos);
        if (ch == '+' || ch == '-') {
            chomping = ch == '+' ? YamlChomping.Keep : YamlChomping.Strip;
            Advance(1);
            ch = At(c.Pos);
            if (ch >= '0' && ch <= '9') {
                increment = ch - '0';
                Advance(1);
            }
        } else if (ch >= '0' && ch <= '9') {
            increment = ch - '0';
            Advance(1);
            ch = At(c.Pos);
            if (ch == '+' || ch == '-') {
                chomping = ch == '+' ? YamlChomping.Keep : YamlChomping.Strip;
                Advance(1);
            }
        }
        if (increment == 0 && At(c.Pos - 1) == '0') {
            throw Error(YamlErrorCodes.InvalidBlockScalar, "A block scalar indentation indicator must be between 1 and 9.");
        }
        char after = At(c.Pos);
        if (!YamlChars.IsBlankOrBreakOrNul(after) || IsNulCharacter(c.Pos)) {
            throw Error(YamlErrorCodes.InvalidBlockScalar, "A block scalar header can contain only the indentation and chomping indicators and a comment.");
        }
        SkipToLineEnd("a block scalar header", YamlErrorCodes.InvalidBlockScalar);

        int headerEnd = c.Pos;
        int headerEndLine = c.Line, headerEndColumn = c.Column;
        if (!HasChar(c.Pos)) {
            return BlockScalarToken(style, chomping, start, startLine, startColumn, headerEnd, headerEnd, Math.Max(c.Indent + 1, 0), headerEndLine, headerEndColumn);
        }
        AdvanceBreak();
        int contentStart = c.Pos;
        int minIndent = Math.Max(c.Indent + 1, 0);
        int indent;

        if (increment > 0) {
            indent = (c.Indent >= 0 ? c.Indent : 0) + increment;
        } else {
            // Auto-detect from the first non-empty line; leading empty lines must not be more indented.
            int p = contentStart;
            int maxEmpty = 0;
            indent = -1;
            while (true) {
                int spaces = 0;
                while (At(p) == ' ') { p++; spaces++; }
                char next = At(p);
                if (YamlChars.IsBreak(next)) {
                    if (spaces > maxEmpty) maxEmpty = spaces;
                    p += next == '\r' && At(p + 1) == '\n' ? 2 : 1;
                    continue;
                }
                if (next == '\0' && !HasChar(p)) {
                    if (spaces > maxEmpty) maxEmpty = spaces;
                    break;
                }
                if (next == '\t' && spaces < minIndent) {
                    throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot be used to indent a block scalar.");
                }
                indent = spaces;
                break;
            }
            if (indent < 0 || indent < minIndent) {
                indent = Math.Max(maxEmpty, minIndent);
            } else if (maxEmpty > indent) {
                throw Error(YamlErrorCodes.InvalidBlockScalar, "A leading empty line of a block scalar is indented more than its first content line.");
            }
        }

        // Content lines.
        int lastContentEnd = -1, lastContentLine = 0, lastContentColumn = 0;
        while (true) {
            int lineStart = c.Pos;
            int spaces = 0;
            while (spaces < indent && At(c.Pos) == ' ') { Advance(1); spaces++; }
            char next = At(c.Pos);
            bool atEnd = next == '\0' && !HasChar(c.Pos);
            // A final line of spaces without a line break stays consumed: it counts as an empty line.
            if (atEnd) break;
            if (spaces < indent) {
                if (YamlChars.IsBreak(next)) {
                    AdvanceBreak();
                    continue;
                }
                c.Pos = lineStart;
                c.Column = 0;
                break;
            }
            if (indent == 0 && IsDocumentMarkerAt(lineStart)) {
                c.Pos = lineStart;
                c.Column = 0;
                break;
            }
            if (YamlChars.IsBreak(next)) {
                AdvanceBreak();
                continue;
            }
            // A content line: scan to its end.
            while (true) {
                int index = YamlChars.IndexOfLineEnd(_text.Slice(c.Pos));
                if (index < 0) {
                    Advance(_text.Length - c.Pos);
                    if (Fill()) continue;
                    break;
                }
                Advance(index);
                if (YamlChars.IsBreak(_text[c.Pos])) break;
                throw Error(YamlErrorCodes.InvalidCharacter, "Found a control character in a block scalar: '" + Describe(_text[c.Pos]) + "'.");
            }
            if (HasChar(c.Pos)) AdvanceBreak();
            lastContentEnd = c.Pos;
            lastContentLine = c.Line;
            lastContentColumn = c.Column;
            if (c.Pos - start > c.MaxScalarLength) CheckScalarLength(start, startLine, startColumn);
        }

        if (chomping != YamlChomping.Keep) {
            int end = lastContentEnd >= 0 ? lastContentEnd : contentStart;
            if (lastContentEnd < 0) {
                // No content: the token ends after the header line break.
                RewindTo(contentStart, headerEndLine + 1, 0);
            } else {
                RewindTo(lastContentEnd, lastContentLine, lastContentColumn);
            }
            return BlockScalarToken(style, chomping, start, startLine, startColumn, contentStart, end, indent, c.Line, c.Column);
        }
        return BlockScalarToken(style, chomping, start, startLine, startColumn, contentStart, c.Pos, indent, c.Line, c.Column);
    }

    private void RewindTo(int pos, int line, int column) {
        _c.Pos = pos;
        _c.Line = line;
        _c.Column = column;
    }

    private YamlToken BlockScalarToken(YamlStyle style, YamlChomping chomping, int start, int startLine, int startColumn, int valueStart, int valueEnd, int indent, int endLine, int endColumn) {
        CheckScalarLength(start, startLine, startColumn);
        YamlToken token = NewToken(YamlTokenKind.Scalar, start, startLine, startColumn);
        token.EndLine = endLine;
        token.EndColumn = endColumn;
        token.Style = style;
        token.Chomping = chomping;
        token.Flags = YamlTokenFlags.MultiLine;
        token.ValueStart = Absolute(valueStart);
        token.ValueEnd = Absolute(valueEnd);
        token.Extra = indent;
        return token;
    }

    /// <summary>Scans a <c>%YAML</c>, <c>%TAG</c> or reserved directive line.</summary>
    private void FetchDirective() {
        YamlScanContext c = _c;
        UnrollIndent(-1);
        RemoveSimpleKey();
        c.SimpleKeyAllowed = false;
        int start = c.Pos, line = c.Line;
        Advance(1);
        int nameStart = c.Pos;
        while (!YamlChars.IsBlankOrBreakOrNul(At(c.Pos))) Advance(1);
        ReadOnlySpan<char> name = _text.Slice(nameStart, c.Pos - nameStart);
        if (name.Length == 0) throw Error(YamlErrorCodes.InvalidDirective, "A directive needs a name after '%'.");
        YamlToken token;
        if (name.SequenceEqual("YAML".AsSpan())) {
            RequireSeparation();
            int major = ScanVersionNumber();
            if (At(c.Pos) != '.') throw Error(YamlErrorCodes.InvalidDirective, "A %YAML directive needs a version such as 1.2.");
            Advance(1);
            int minor = ScanVersionNumber();
            token = NewToken(YamlTokenKind.VersionDirective, start, line, 0);
            token.Extra = major;
            token.Extra2 = minor;
        } else if (name.SequenceEqual("TAG".AsSpan())) {
            RequireSeparation();
            int handleStart = c.Pos;
            if (At(c.Pos) != '!') throw Error(YamlErrorCodes.InvalidDirective, "A %TAG handle must start with '!'.");
            Advance(1);
            if (!YamlChars.IsBlank(At(c.Pos))) {
                while (YamlChars.IsWord(At(c.Pos))) Advance(1);
                if (At(c.Pos) != '!') throw Error(YamlErrorCodes.InvalidDirective, "A named %TAG handle must end with '!'.");
                Advance(1);
            }
            int handleEnd = c.Pos;
            RequireSeparation();
            int prefixStart = c.Pos;
            char first = At(c.Pos);
            if (first != '!' && !YamlChars.IsTagChar(first) && first != '%') throw Error(YamlErrorCodes.InvalidDirective, "A %TAG prefix must be a local tag or a URI.");
            while (true) {
                char ch = At(c.Pos);
                if (ch == '%') { ScanPercentEscape(); continue; }
                if (!YamlChars.IsUri(ch)) break;
                Advance(1);
            }
            token = NewToken(YamlTokenKind.TagDirective, start, line, 0);
            token.ValueStart = Absolute(handleStart);
            token.ValueEnd = Absolute(handleEnd);
            token.Extra = Absolute(prefixStart);
            token.Extra2 = Absolute(c.Pos);
        } else {
            // Reserved directive: parameters up to an optional comment.
            int end = c.Pos;
            while (true) {
                while (YamlChars.IsBlank(At(c.Pos))) Advance(1);
                char ch = At(c.Pos);
                if (ch == '#' || YamlChars.IsBreak(ch) || (ch == '\0' && !HasChar(c.Pos))) break;
                while (!YamlChars.IsBlankOrBreakOrNul(At(c.Pos))) Advance(1);
                end = c.Pos;
            }
            RewindTo(end, line, end - start);
            token = NewToken(YamlTokenKind.ReservedDirective, start, line, 0);
            token.ValueStart = Absolute(nameStart);
            token.ValueEnd = Absolute(nameStart + name.Length);
        }
        char after = At(c.Pos);
        if (!YamlChars.IsBlankOrBreakOrNul(after)) throw Error(YamlErrorCodes.InvalidDirective, "Unexpected content in a directive.");
        Finish(ref token);
        SkipToLineEnd("a directive", YamlErrorCodes.InvalidDirective);
    }

    private void RequireSeparation() {
        if (!YamlChars.IsBlank(At(_c.Pos))) throw Error(YamlErrorCodes.InvalidDirective, "Directive parameters must be separated by whitespace.");
        while (YamlChars.IsBlank(At(_c.Pos))) Advance(1);
    }

    private int ScanVersionNumber() {
        int value = 0;
        int digits = 0;
        while (true) {
            char ch = At(_c.Pos);
            if (ch < '0' || ch > '9') break;
            if (value < 100000) value = value * 10 + (ch - '0');
            digits++;
            Advance(1);
        }
        if (digits == 0) throw Error(YamlErrorCodes.InvalidDirective, "A %YAML directive needs a version such as 1.2.");
        return value;
    }
}
