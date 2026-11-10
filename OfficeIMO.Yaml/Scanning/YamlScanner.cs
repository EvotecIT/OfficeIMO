using System.Runtime.CompilerServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// The YAML 1.2.2 scanner. It turns the decoded text into <see cref="YamlToken"/> values without allocating.
/// Whitespace, line breaks and comments are not tokens: they are the gaps between token spans.
/// </summary>
/// <remarks>
/// The scanner is a ref struct over the current text span so string input is never copied. Everything that must
/// survive between reads (positions, indentation and flow stacks, implicit-key candidates and the token queue) lives
/// in <see cref="YamlScanContext"/>. Buffered input grows the span in place while a token is scanned; it is only
/// compacted between tokens.
/// </remarks>
internal ref partial struct YamlScanner {
    private const int MaxImplicitKeyLength = 1024;

    private ReadOnlySpan<char> _text;
    private readonly YamlScanContext _c;

    public YamlScanner(ReadOnlySpan<char> text, YamlScanContext context) {
        _text = text;
        _c = context;
    }

    /// <summary>Gets the current text span. Absolute offsets map to it by subtracting <see cref="YamlScanContext.Base"/>.</summary>
    public readonly ReadOnlySpan<char> Text => _text;

    // ---------------------------------------------------------------- token queue

    /// <summary>Returns the next token, scanning ahead as far as implicit keys require.</summary>
    public ref YamlToken Peek() {
        YamlScanContext c = _c;
        while (NeedMoreTokens()) FetchNextToken();
        return ref c.Tokens[c.TokenHead];
    }

    /// <summary>Consumes the token returned by <see cref="Peek"/>.</summary>
    public void Skip() {
        YamlScanContext c = _c;
        c.TokenHead++;
        c.TokenCount--;
        c.TokensParsed++;
        if (c.TokenCount == 0) c.TokenHead = 0;
    }

    private bool NeedMoreTokens() {
        YamlScanContext c = _c;
        if (c.TokenCount == 0) return !c.StreamEndProduced;
        if (c.StreamEndProduced) return false;
        StaleSimpleKeys();
        YamlSimpleKey[] keys = c.SimpleKeys;
        for (int i = 0; i <= c.FlowLevel; i++) {
            if (keys[i].Possible && keys[i].TokenNumber == c.TokensParsed) return true;
        }
        return false;
    }

    private void Enqueue(in YamlToken token) {
        YamlScanContext c = _c;
        EnsureTokenCapacity();
        c.Tokens[c.TokenHead + c.TokenCount] = token;
        c.TokenCount++;
    }

    private void Insert(int tokenNumber, in YamlToken token) {
        YamlScanContext c = _c;
        EnsureTokenCapacity();
        int index = c.TokenHead + (tokenNumber - c.TokensParsed);
        int tail = c.TokenHead + c.TokenCount;
        if (index < tail) Array.Copy(c.Tokens, index, c.Tokens, index + 1, tail - index);
        c.Tokens[index] = token;
        c.TokenCount++;
    }

    private void EnsureTokenCapacity() {
        YamlScanContext c = _c;
        if (c.TokenHead + c.TokenCount < c.Tokens.Length) return;
        if (c.TokenHead > 0) {
            Array.Copy(c.Tokens, c.TokenHead, c.Tokens, 0, c.TokenCount);
            c.TokenHead = 0;
            if (c.TokenCount < c.Tokens.Length) return;
        }
        c.Tokens = YamlScanContext.Grow(c.Tokens, c.TokenCount);
    }

    /// <summary>Records the end of a token that has a source span, for block ends, empty nodes and line tracking.</summary>
    private void Finish(ref YamlToken token) {
        YamlScanContext c = _c;
        c.LastEnd = token.End;
        c.LastEndLine = token.EndLine;
        c.LastEndColumn = token.EndColumn;
        c.LastLine = token.EndColumn == 0 && token.EndLine > token.StartLine ? token.EndLine - 1 : token.EndLine;
        Enqueue(token);
    }

    // ---------------------------------------------------------------- character access

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private char At(int index) => (uint)index < (uint)_text.Length ? _text[index] : AtSlow(index);

    private char AtSlow(int index) {
        if (index < 0) return '\0';
        while (index >= _text.Length && Fill()) { }
        return index < _text.Length ? _text[index] : '\0';
    }

    /// <summary>Returns whether a character exists at the index, reading more buffered input if needed.</summary>
    private bool HasChar(int index) {
        while (index >= _text.Length) {
            if (!Fill()) return false;
        }
        return true;
    }

    private bool Fill() {
        YamlScanContext c = _c;
        if (c.IsFinal) return false;
        bool read = c.Fill();
        _text = new ReadOnlySpan<char>(c.Buffer, 0, c.Length);
        return read;
    }

    private void CompactAtSafePoint() {
        YamlScanContext c = _c;
        if (c.Source is null) return;
        int before = c.Base;
        c.Compact();
        if (c.Base != before) _text = new ReadOnlySpan<char>(c.Buffer, 0, c.Length);
    }

    private readonly int Absolute(int relative) => relative + _c.Base;

    /// <summary>Advances over characters that contain no line break.</summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void Advance(int count) {
        _c.Pos += count;
        _c.Column += count;
    }

    /// <summary>Advances over one line break (<c>\r\n</c>, <c>\n</c> or <c>\r</c>) at the current position.</summary>
    private void AdvanceBreak() {
        YamlScanContext c = _c;
        if (At(c.Pos) == '\r' && At(c.Pos + 1) == '\n') c.Pos += 2;
        else c.Pos += 1;
        c.Line++;
        c.Column = 0;
    }

    private bool IsDocumentMarker(int index, char marker) =>
        At(index) == marker && At(index + 1) == marker && At(index + 2) == marker && YamlChars.IsBlankOrBreakOrNul(At(index + 3)) && !IsNulCharacter(index + 3);

    private bool IsDocumentMarkerAt(int index) {
        char ch = At(index);
        return (ch == '-' || ch == '.') && IsDocumentMarker(index, ch);
    }

    /// <summary>True when the character at the index is a real U+0000 rather than the end of input.</summary>
    private bool IsNulCharacter(int index) => At(index) == '\0' && HasChar(index);

    private YamlToken NewToken(YamlTokenKind kind, int start, int startLine, int startColumn) {
        YamlScanContext c = _c;
        return new YamlToken {
            Kind = kind,
            Start = Absolute(start),
            StartLine = startLine,
            StartColumn = startColumn,
            End = Absolute(c.Pos),
            EndLine = c.Line,
            EndColumn = c.Column
        };
    }

    private YamlToken ZeroWidth(YamlTokenKind kind, int offset, int line, int column) => new YamlToken {
        Kind = kind,
        Start = offset,
        End = offset,
        StartLine = line,
        StartColumn = column,
        EndLine = line,
        EndColumn = column
    };

    // ---------------------------------------------------------------- errors

    private readonly YamlException Error(string code, string message) {
        YamlScanContext c = _c;
        return new YamlException(code, message, new YamlPosition((long)c.Base + c.Pos, c.Line + 1, c.Column + 1));
    }

    private static YamlException Error(string code, string message, int offset, int line, int column) =>
        new YamlException(code, message, new YamlPosition(offset, line + 1, column + 1));

    private void CheckScalarLength(int startRelative, int startLine, int startColumn) {
        YamlScanContext c = _c;
        if (c.Pos - startRelative > c.MaxScalarLength) {
            throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxScalarLength), c.MaxScalarLength, new YamlPosition(Absolute(startRelative), startLine + 1, startColumn + 1));
        }
    }

    // ---------------------------------------------------------------- dispatch

    private void FetchNextToken() {
        YamlScanContext c = _c;
        if (!c.StreamStartProduced) {
            FetchStreamStart();
            return;
        }
        if (--c.CancelCountdown <= 0) {
            c.CancelCountdown = YamlScanContext.CancelInterval;
            c.Cancellation.ThrowIfCancellationRequested();
        }

        ScanToNextToken();
        StaleSimpleKeys();
        UnrollIndent(c.Column);

        int p = c.Pos;
        char ch = At(p);
        if (ch == '\0' && !HasChar(p)) {
            FetchStreamEnd();
            return;
        }

        bool firstOnLine = c.LastLine != c.Line;
        if (c.FlowLevel > 0 && firstOnLine && c.Column <= c.Indent) {
            throw Error(YamlErrorCodes.InvalidIndentation, "Flow content must be indented more than the enclosing block collection.");
        }

        if (c.Column == 0) {
            if (ch == '%') {
                if (c.InDocument) throw Error(YamlErrorCodes.InvalidDocument, "A directive must follow a document end marker ('...') or start the stream.");
                FetchDirective();
                return;
            }
            if (ch == '-' && IsDocumentMarker(p, '-')) {
                FetchDocumentIndicator(YamlTokenKind.DocumentStart);
                return;
            }
            if (ch == '.' && IsDocumentMarker(p, '.')) {
                FetchDocumentIndicator(YamlTokenKind.DocumentEnd);
                return;
            }
        }

        c.InDocument = true;
        char next = At(p + 1);
        bool flow = c.FlowLevel > 0;
        switch (ch) {
            case '[':
                FetchFlowCollectionStart(YamlTokenKind.FlowSequenceStart);
                return;
            case '{':
                FetchFlowCollectionStart(YamlTokenKind.FlowMappingStart);
                return;
            case ']':
                FetchFlowCollectionEnd(YamlTokenKind.FlowSequenceEnd);
                return;
            case '}':
                FetchFlowCollectionEnd(YamlTokenKind.FlowMappingEnd);
                return;
            case ',':
                FetchFlowEntry();
                return;
            case '-':
                if (YamlChars.IsBlankOrBreakOrNul(next)) {
                    FetchBlockEntry();
                    return;
                }
                break;
            case '?':
                if (YamlChars.IsBlankOrBreakOrNul(next)) {
                    FetchKey();
                    return;
                }
                break;
            case ':':
                if (YamlChars.IsBlankOrBreakOrNul(next) || (flow && (YamlChars.IsFlowIndicator(next) || c.AdjacentValue))) {
                    FetchValue();
                    return;
                }
                break;
            case '*':
                FetchAnchor(YamlTokenKind.Alias);
                return;
            case '&':
                FetchAnchor(YamlTokenKind.Anchor);
                return;
            case '!':
                FetchTag();
                return;
            case '|':
            case '>':
                if (!flow) {
                    FetchBlockScalar(ch == '|' ? YamlStyle.Literal : YamlStyle.Folded);
                    return;
                }
                break;
            case '\'':
                FetchFlowScalar(single: true);
                return;
            case '"':
                FetchFlowScalar(single: false);
                return;
        }

        if (CanStartPlain(ch, next, flow)) {
            FetchPlainScalar();
            return;
        }
        if (ch == '%' ) throw Error(YamlErrorCodes.InvalidDocument, "A directive must start at the beginning of a line, outside a document.");
        if (ch == '#') throw Error(YamlErrorCodes.InvalidCharacter, "A comment must be separated from the preceding token by whitespace.");
        if (ch == '\t') throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot be used here.");
        throw Error(YamlErrorCodes.InvalidCharacter, "Found a character that cannot start any token: '" + Describe(ch) + "'.");
    }

    private static string Describe(char ch) =>
        ch < 0x20 || ch == 0x7F || ch is '￾' or '￿' or '﻿' ? "\\u" + ((int)ch).ToString("X4", CultureInfo.InvariantCulture) : ch.ToString();

    private static bool CanStartPlain(char ch, char next, bool flow) {
        if (YamlChars.IsBlankOrBreakOrNul(ch)) return false;
        if (ch == '-' || ch == '?' || ch == ':') {
            return !YamlChars.IsBlankOrBreakOrNul(next) && !(flow && YamlChars.IsFlowIndicator(next));
        }
        if (YamlChars.IsIndicator(ch)) return false;
        return !YamlChars.IsNonPrintable(ch) && ch != '﻿';
    }

    // ---------------------------------------------------------------- whitespace, comments, line breaks

    /// <summary>
    /// Skips whitespace, comments and line breaks before the next token and applies the tab rules: a tab may not
    /// stand in for indentation the enclosing block structure requires, and a tab before a block indicator or an
    /// implicit key that opens a block collection is rejected later through <see cref="YamlScanContext.TabBefore"/>.
    /// </summary>
    private void ScanToNextToken() {
        YamlScanContext c = _c;
        int lines = 0;
        while (true) {
            if (c.Column == 0 && !c.InDocument && At(c.Pos) == '﻿') {
                c.Pos++;
            }

            int firstTabColumn = -1;
            int runStart = c.Pos;
            while (true) {
                char ch = At(c.Pos);
                if (ch == ' ') { Advance(1); continue; }
                if (ch == '\t') {
                    if (firstTabColumn < 0) firstTabColumn = c.Column;
                    Advance(1);
                    continue;
                }
                break;
            }

            char current = At(c.Pos);
            // A line start separates a comment like whitespace does; only a block scalar ends exactly there.
            if (current == '#' && (c.Pos > runStart || Absolute(c.Pos) != c.LastEnd || c.Column == 0)) {
                SkipComment();
                current = At(c.Pos);
            }

            if (YamlChars.IsBreak(current)) {
                AdvanceBreak();
                if (c.FlowLevel == 0) c.SimpleKeyAllowed = true;
                if (++lines >= 64) {
                    lines = 0;
                    if (--c.CancelCountdown <= 0) {
                        c.CancelCountdown = YamlScanContext.CancelInterval;
                        c.Cancellation.ThrowIfCancellationRequested();
                    }
                    CompactAtSafePoint();
                }
                continue;
            }

            c.TabBefore = false;
            if (firstTabColumn >= 0 && !(current == '\0' && !HasChar(c.Pos))) {
                bool firstOnLine = c.LastLine != c.Line;
                if (firstOnLine && firstTabColumn <= c.Indent) {
                    throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot be used for indentation.", Absolute(c.Pos - (c.Column - firstTabColumn)), c.Line, firstTabColumn);
                }
                c.TabBefore = true;
            }
            return;
        }
    }

    /// <summary>Skips a comment that starts at the current '#', recording it when comment events are requested.</summary>
    private void SkipComment() {
        YamlScanContext c = _c;
        int start = c.Pos;
        int startColumn = c.Column;
        while (true) {
            int index = YamlChars.IndexOfLineEnd(_text.Slice(c.Pos));
            if (index < 0) {
                Advance(_text.Length - c.Pos);
                if (Fill()) continue;
                break;
            }
            Advance(index);
            char ch = _text[c.Pos];
            if (YamlChars.IsBreak(ch)) break;
            throw Error(YamlErrorCodes.InvalidCharacter, "Comments cannot contain control characters.");
        }
        if (c.IncludeComments) {
            if (c.CommentHead + c.CommentCount >= c.Comments.Length) {
                if (c.CommentHead > 0) {
                    Array.Copy(c.Comments, c.CommentHead, c.Comments, 0, c.CommentCount);
                    c.CommentHead = 0;
                }
                if (c.CommentCount >= c.Comments.Length) c.Comments = YamlScanContext.Grow(c.Comments, c.CommentCount);
            }
            c.Comments[c.CommentHead + c.CommentCount] = new YamlComment {
                Start = Absolute(start),
                End = Absolute(c.Pos),
                Line = c.Line,
                Column = startColumn
            };
            c.CommentCount++;
        }
    }

    /// <summary>Skips blanks and an optional comment to the end of the line; anything else is an error.</summary>
    private void SkipToLineEnd(string context, string code) {
        YamlScanContext c = _c;
        int before = c.Pos;
        while (YamlChars.IsBlank(At(c.Pos))) Advance(1);
        char ch = At(c.Pos);
        if (ch == '#') {
            if (c.Pos == before) throw Error(YamlErrorCodes.InvalidCharacter, "A comment must be separated from " + context + " by whitespace.");
            SkipComment();
            ch = At(c.Pos);
        }
        if (YamlChars.IsBreak(ch) || (ch == '\0' && !HasChar(c.Pos))) return;
        throw Error(code, "Unexpected content after " + context + ".");
    }

    // ---------------------------------------------------------------- implicit keys and indentation

    private void StaleSimpleKeys() {
        YamlScanContext c = _c;
        YamlSimpleKey[] keys = c.SimpleKeys;
        int absolute = Absolute(c.Pos);
        for (int i = 0; i <= c.FlowLevel; i++) {
            ref YamlSimpleKey key = ref keys[i];
            if (key.Possible && (key.Line < c.Line || absolute - key.Offset > MaxImplicitKeyLength)) {
                if (key.Required) {
                    throw Error(YamlErrorCodes.InvalidKey, "Could not find the ':' that completes this implicit key.", key.Offset, key.Line, key.Column);
                }
                key.Possible = false;
            }
        }
    }

    private void SaveSimpleKey() {
        YamlScanContext c = _c;
        if (!c.SimpleKeyAllowed) return;
        RemoveSimpleKey();
        ref YamlSimpleKey key = ref c.SimpleKeys[c.FlowLevel];
        key.Possible = true;
        key.Required = c.FlowLevel == 0 && c.Indent == c.Column;
        key.TabBefore = c.TabBefore;
        key.TokenNumber = c.TokensParsed + c.TokenCount;
        key.Offset = Absolute(c.Pos);
        key.Line = c.Line;
        key.Column = c.Column;
    }

    private void RemoveSimpleKey() {
        YamlScanContext c = _c;
        ref YamlSimpleKey key = ref c.SimpleKeys[c.FlowLevel];
        if (key.Possible && key.Required) {
            throw Error(YamlErrorCodes.InvalidKey, "Could not find the ':' that completes this implicit key.", key.Offset, key.Line, key.Column);
        }
        key.Possible = false;
    }

    private void IncreaseFlowLevel() {
        YamlScanContext c = _c;
        if (c.FlowLevel + c.IndentCount >= c.MaxDepth) ThrowDepth();
        c.FlowLevel++;
        if (c.FlowLevel >= c.SimpleKeys.Length) c.SimpleKeys = YamlScanContext.Grow(c.SimpleKeys, c.FlowLevel);
        c.SimpleKeys[c.FlowLevel] = default;
    }

    private void DecreaseFlowLevel() {
        YamlScanContext c = _c;
        if (c.FlowLevel > 0) c.FlowLevel--;
    }

    private readonly void ThrowDepth() {
        YamlScanContext c = _c;
        throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxDepth), c.MaxDepth, new YamlPosition((long)c.Base + c.Pos, c.Line + 1, c.Column + 1));
    }

    /// <summary>Opens a block collection when the column is deeper than the current indentation.</summary>
    private bool RollIndent(int column, int tokenNumber, YamlTokenKind kind, int offset, int line) {
        YamlScanContext c = _c;
        if (c.FlowLevel > 0 || c.Indent >= column) return false;
        if (c.FlowLevel + c.IndentCount + 1 >= c.MaxDepth) ThrowDepth();
        if (c.IndentCount >= c.Indents.Length) c.Indents = YamlScanContext.Grow(c.Indents, c.IndentCount);
        c.Indents[c.IndentCount++] = c.Indent;
        c.Indent = column;
        YamlToken token = ZeroWidth(kind, offset, line, column);
        if (tokenNumber < 0) Enqueue(token);
        else Insert(tokenNumber, token);
        return true;
    }

    /// <summary>Closes block collections deeper than the column. Block ends sit at the end of the last content.</summary>
    private void UnrollIndent(int column) {
        YamlScanContext c = _c;
        if (c.FlowLevel > 0) return;
        while (c.Indent > column) {
            Enqueue(ZeroWidth(YamlTokenKind.BlockEnd, c.LastEnd, c.LastEndLine, c.LastEndColumn));
            c.Indent = c.Indents[--c.IndentCount];
        }
    }

    // ---------------------------------------------------------------- simple tokens

    private void FetchStreamStart() {
        YamlScanContext c = _c;
        c.Indent = -1;
        c.SimpleKeyAllowed = true;
        c.SimpleKeys[0] = default;
        c.StreamStartProduced = true;
        Enqueue(ZeroWidth(YamlTokenKind.StreamStart, Absolute(c.Pos), 0, 0));
    }

    private void FetchStreamEnd() {
        YamlScanContext c = _c;
        UnrollIndent(-1);
        RemoveSimpleKey();
        c.SimpleKeyAllowed = false;
        c.StreamEndProduced = true;
        Enqueue(ZeroWidth(YamlTokenKind.StreamEnd, Absolute(c.Pos), c.Line, c.Column));
    }

    private void FetchDocumentIndicator(YamlTokenKind kind) {
        YamlScanContext c = _c;
        UnrollIndent(-1);
        RemoveSimpleKey();
        c.SimpleKeyAllowed = false;
        c.AdjacentValue = false;
        int start = c.Pos;
        int line = c.Line;
        Advance(3);
        YamlToken token = NewToken(kind, start, line, 0);
        if (kind == YamlTokenKind.DocumentEnd) {
            c.InDocument = false;
            Finish(ref token);
            SkipToLineEnd("a document end marker", YamlErrorCodes.InvalidDocument);
        } else {
            c.InDocument = true;
            Finish(ref token);
        }
    }

    private void FetchFlowCollectionStart(YamlTokenKind kind) {
        YamlScanContext c = _c;
        SaveSimpleKey();
        IncreaseFlowLevel();
        c.SimpleKeyAllowed = true;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(kind, start, line, column);
        Finish(ref token);
    }

    private void FetchFlowCollectionEnd(YamlTokenKind kind) {
        YamlScanContext c = _c;
        RemoveSimpleKey();
        DecreaseFlowLevel();
        c.SimpleKeyAllowed = false;
        c.AdjacentValue = true;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(kind, start, line, column);
        Finish(ref token);
    }

    private void FetchFlowEntry() {
        YamlScanContext c = _c;
        RemoveSimpleKey();
        c.SimpleKeyAllowed = true;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(YamlTokenKind.FlowEntry, start, line, column);
        Finish(ref token);
    }

    private void FetchBlockEntry() {
        YamlScanContext c = _c;
        if (c.FlowLevel > 0) throw Error(YamlErrorCodes.UnexpectedToken, "A block sequence entry ('- ') cannot appear inside a flow collection.");
        if (!c.SimpleKeyAllowed) throw Error(YamlErrorCodes.UnexpectedToken, "A block sequence entry is not allowed here; block collections start on a new line.");
        if (c.TabBefore) throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot separate block collection indicators.");
        RollIndent(c.Column, -1, YamlTokenKind.BlockSequenceStart, Absolute(c.Pos), c.Line);
        RemoveSimpleKey();
        c.SimpleKeyAllowed = true;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(YamlTokenKind.BlockEntry, start, line, column);
        Finish(ref token);
    }

    private void FetchKey() {
        YamlScanContext c = _c;
        if (c.FlowLevel == 0) {
            if (!c.SimpleKeyAllowed) throw Error(YamlErrorCodes.UnexpectedToken, "An explicit key ('? ') is not allowed here; block collections start on a new line.");
            if (c.TabBefore) throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot separate block collection indicators.");
            RollIndent(c.Column, -1, YamlTokenKind.BlockMappingStart, Absolute(c.Pos), c.Line);
        }
        RemoveSimpleKey();
        c.SimpleKeyAllowed = c.FlowLevel == 0;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(YamlTokenKind.Key, start, line, column);
        Finish(ref token);
    }

    private void FetchValue() {
        YamlScanContext c = _c;
        ref YamlSimpleKey key = ref c.SimpleKeys[c.FlowLevel];
        if (key.Possible) {
            Insert(key.TokenNumber, ZeroWidth(YamlTokenKind.Key, key.Offset, key.Line, key.Column));
            bool opened = RollIndent(key.Column, key.TokenNumber, YamlTokenKind.BlockMappingStart, key.Offset, key.Line);
            if (opened && key.TabBefore) {
                throw Error(YamlErrorCodes.InvalidIndentation, "A tab character cannot indent a block mapping.", key.Offset, key.Line, key.Column);
            }
            key.Possible = false;
            c.SimpleKeyAllowed = false;
        } else {
            if (c.FlowLevel == 0) {
                if (!c.SimpleKeyAllowed) throw Error(YamlErrorCodes.InvalidValue, "A mapping value is not allowed here.");
                RollIndent(c.Column, -1, YamlTokenKind.BlockMappingStart, Absolute(c.Pos), c.Line);
            }
            c.SimpleKeyAllowed = c.FlowLevel == 0;
        }
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        YamlToken token = NewToken(YamlTokenKind.Value, start, line, column);
        Finish(ref token);
    }

    private void FetchAnchor(YamlTokenKind kind) {
        YamlScanContext c = _c;
        SaveSimpleKey();
        c.SimpleKeyAllowed = false;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        Advance(1);
        int nameStart = c.Pos;
        while (true) {
            char ch = At(c.Pos);
            if (YamlChars.IsBlankOrBreakOrNul(ch) || YamlChars.IsFlowIndicator(ch) || YamlChars.IsNonPrintable(ch) || ch == '﻿') break;
            Advance(1);
        }
        if (c.Pos == nameStart) {
            throw Error(YamlErrorCodes.InvalidAnchor, kind == YamlTokenKind.Alias ? "An alias needs an anchor name after '*'." : "An anchor needs a name after '&'.");
        }
        YamlToken token = NewToken(kind, start, line, column);
        token.ValueStart = Absolute(nameStart);
        token.ValueEnd = Absolute(c.Pos);
        Finish(ref token);
    }

    private void FetchTag() {
        YamlScanContext c = _c;
        SaveSimpleKey();
        c.SimpleKeyAllowed = false;
        c.AdjacentValue = false;
        int start = c.Pos, line = c.Line, column = c.Column;
        YamlTokenFlags flags = YamlTokenFlags.None;
        int suffixStart;
        int suffixEnd;
        int handleEnd;
        if (At(start + 1) == '<') {
            Advance(2);
            suffixStart = c.Pos;
            while (true) {
                char ch = At(c.Pos);
                if (ch == '%') { ScanPercentEscape(); flags |= YamlTokenFlags.PercentEscapes; continue; }
                if (!YamlChars.IsUri(ch)) break;
                Advance(1);
            }
            suffixEnd = c.Pos;
            if (At(c.Pos) != '>' || suffixEnd == suffixStart) throw Error(YamlErrorCodes.InvalidTag, "A verbatim tag must be a non-empty URI closed by '>'.");
            Advance(1);
            flags |= YamlTokenFlags.Verbatim;
            handleEnd = start;
        } else {
            int p = start + 1;
            while (YamlChars.IsWord(At(p))) p++;
            if (At(p) == '!') {
                handleEnd = p + 1;
            } else {
                handleEnd = start + 1;
            }
            Advance(handleEnd - start);
            suffixStart = c.Pos;
            while (true) {
                char ch = At(c.Pos);
                if (ch == '%') { ScanPercentEscape(); flags |= YamlTokenFlags.PercentEscapes; continue; }
                if (!YamlChars.IsTagChar(ch)) break;
                Advance(1);
            }
            suffixEnd = c.Pos;
            if (suffixEnd == suffixStart && handleEnd - start > 1) {
                throw Error(YamlErrorCodes.InvalidTag, "A tag with a named or secondary handle needs a suffix.");
            }
        }
        char after = At(c.Pos);
        if (!YamlChars.IsBlankOrBreakOrNul(after) && !(c.FlowLevel > 0 && YamlChars.IsFlowIndicator(after))) {
            throw Error(YamlErrorCodes.InvalidTag, "A tag must be followed by whitespace.");
        }
        YamlToken token = NewToken(YamlTokenKind.Tag, start, line, column);
        token.Flags = flags;
        token.ValueStart = Absolute(suffixStart);
        token.ValueEnd = Absolute(suffixEnd);
        token.Extra = Absolute(handleEnd);
        Finish(ref token);
    }

    private void ScanPercentEscape() {
        if (YamlChars.HexValue(At(_c.Pos + 1)) < 0 || YamlChars.HexValue(At(_c.Pos + 2)) < 0) {
            throw Error(YamlErrorCodes.InvalidTag, "A '%' in a tag must be followed by two hexadecimal digits.");
        }
        Advance(3);
    }

    private void FetchPlainScalar() {
        YamlScanContext c = _c;
        SaveSimpleKey();
        c.SimpleKeyAllowed = false;
        c.AdjacentValue = false;
        YamlToken token = ScanPlainScalar();
        Finish(ref token);
    }

    private void FetchFlowScalar(bool single) {
        YamlScanContext c = _c;
        SaveSimpleKey();
        c.SimpleKeyAllowed = false;
        YamlToken token = ScanFlowScalar(single);
        c.AdjacentValue = true;
        Finish(ref token);
    }

    private void FetchBlockScalar(YamlStyle style) {
        YamlScanContext c = _c;
        RemoveSimpleKey();
        c.SimpleKeyAllowed = true;
        c.AdjacentValue = false;
        YamlToken token = ScanBlockScalar(style);
        Finish(ref token);
    }
}
