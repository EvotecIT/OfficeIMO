namespace OfficeIMO.Toml;

/// <summary>
/// A forward-only pull reader over TOML text, modeled on <c>Utf8JsonReader</c>. Each <see cref="Read"/> advances to the
/// next header bracket, key segment, value, array or inline-table bracket, or comment, and allocates nothing unless the
/// caller asks for a string.
/// </summary>
/// <remarks>
/// <para>The reader checks TOML 1.1 syntax completely: characters, strings and escapes, numbers and their ranges,
/// date-time fields, and where keys, values, commas and newlines may appear. Table rules that need the whole document,
/// such as duplicate keys and table redefinition, are checked by <see cref="TomlDocument"/>.</para>
/// <para>A key/value pair reads as one <see cref="TomlTokenType.Key"/> token per dotted segment followed by the value;
/// a header reads as its start token, its key segments and its end token.</para>
/// </remarks>
public sealed partial class TomlReader {
    private const int CancellationInterval = 4096;

    private enum State : byte {
        LineStart,
        HeaderKey,
        HeaderAfterKey,
        Key,
        AfterKey,
        Value,
        AfterValue,
        ArrayFirst,
        ArrayItem,
        InlineFirst,
        InlineKey,
        End
    }

    private const byte ArrayContext = 1;
    private const byte InlineContext = 2;

    private readonly ReadOnlyMemory<char> _text;
    private readonly TomlReaderOptions _options;
    private readonly CancellationToken _cancellationToken;
    private byte[] _contexts = new byte[8];
    private int _depth;
    private State _state;
    private bool _headerIsArray;
    private int _pos;
    private int _line = 1;
    private int _lineStart;
    private int _tokenCount;

    private TomlTokenType _tokenType;
    private int _tokenStart;
    private int _tokenEnd;
    private int _startLine;
    private int _startColumn;
    private int _endLine;
    private int _endColumn;
    private TomlStringStyle _style;

    /// <summary>Creates a reader over TOML text.</summary>
    /// <param name="toml">The TOML text.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed at bounded intervals while reading.</param>
    public TomlReader(string toml, TomlReaderOptions? options = null, CancellationToken cancellationToken = default)
        : this((toml ?? throw new ArgumentNullException(nameof(toml))).AsMemory(), options, cancellationToken) {
    }

    /// <summary>Creates a reader over TOML text.</summary>
    /// <param name="toml">The TOML text. It must not change while the reader uses it.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Observed at bounded intervals while reading.</param>
    public TomlReader(ReadOnlyMemory<char> toml, TomlReaderOptions? options = null, CancellationToken cancellationToken = default) {
        _text = toml;
        _options = options ?? new TomlReaderOptions();
        _cancellationToken = cancellationToken;
    }

    /// <summary>Gets the token the reader is on.</summary>
    public TomlTokenType TokenType => _tokenType;

    /// <summary>Gets the nesting of arrays and inline tables the current token is inside; top-level tokens are at depth 0.</summary>
    public int Depth => _tokenType is TomlTokenType.StartArray or TomlTokenType.StartInlineTable ? _depth - 1 : _depth;

    /// <summary>Gets where the current token starts.</summary>
    public TomlPosition Start => new TomlPosition(_tokenStart, _startLine, _startColumn);

    /// <summary>Gets where the current token ends (exclusive).</summary>
    public TomlPosition End => new TomlPosition(_tokenEnd, _endLine, _endColumn);

    /// <summary>
    /// Gets the current token's source text: a key or string with its quotes, a number, date-time or boolean as written,
    /// a bracket or brace, or a comment with its <c>#</c>.
    /// </summary>
    public ReadOnlySpan<char> ValueSpan => _text.Span.Slice(_tokenStart, _tokenEnd - _tokenStart);

    /// <summary>
    /// Gets how the current <see cref="TomlTokenType.Key"/> or <see cref="TomlTokenType.String"/> is quoted:
    /// <see cref="TomlStringStyle.Bare"/> for bare keys.
    /// </summary>
    public TomlStringStyle StringStyle => _style;

    internal int TokenStartOffset => _tokenStart;

    internal int TokenEndOffset => _tokenEnd;

    /// <summary>
    /// Returns the decoded text of a key or string (escapes applied, the first newline of a multi-line string trimmed,
    /// <c>\r\n</c> normalized to <c>\n</c>), or the text of a comment after its <c>#</c>.
    /// </summary>
    /// <exception cref="InvalidOperationException">The token is not a key, string or comment.</exception>
    public string GetString() {
        ReadOnlySpan<char> raw = ValueSpan;
        return _tokenType switch {
            TomlTokenType.Key or TomlTokenType.String => TomlScalars.Decode(raw, _style),
            TomlTokenType.Comment => raw.Slice(1).ToString(),
            _ => throw WrongToken("a key, string or comment")
        };
    }

    /// <summary>Returns the value of an integer token.</summary>
    /// <exception cref="InvalidOperationException">The token is not an integer.</exception>
    public long GetInt64() {
        if (_tokenType != TomlTokenType.Integer) throw WrongToken("an integer");
        TomlScalars.TryParseInteger(ValueSpan, out long value, out _, out _);
        return value;
    }

    /// <summary>Returns the value of a float or integer token.</summary>
    /// <exception cref="InvalidOperationException">The token is not a float or integer.</exception>
    public double GetDouble() {
        if (_tokenType == TomlTokenType.Integer) return GetInt64();
        if (_tokenType != TomlTokenType.Float) throw WrongToken("a float");
        return TomlScalars.ParseFloat(ValueSpan);
    }

    /// <summary>Returns the value of a boolean token.</summary>
    /// <exception cref="InvalidOperationException">The token is not a boolean.</exception>
    public bool GetBoolean() {
        if (_tokenType != TomlTokenType.Boolean) throw WrongToken("a boolean");
        return ValueSpan[0] == 't';
    }

    /// <summary>Returns the value of a date, time or date-time token.</summary>
    /// <exception cref="InvalidOperationException">The token is not a date or time.</exception>
    public TomlDateTime GetDateTime() {
        if (_tokenType is not (TomlTokenType.OffsetDateTime or TomlTokenType.LocalDateTime or TomlTokenType.LocalDate or TomlTokenType.LocalTime)) {
            throw WrongToken("a date or time");
        }
        TomlScalars.TryParseDateTime(ValueSpan, out TomlDateTime value, out _);
        return value;
    }

    private InvalidOperationException WrongToken(string expected) =>
        new InvalidOperationException("The reader is on " + _tokenType + ", not " + expected + ".");

    /// <summary>
    /// Advances to the next token. Returns <see langword="false"/> at the end of the input, after checking that every
    /// array and inline table is closed.
    /// </summary>
    /// <exception cref="TomlException">The input is malformed or exceeds a limit.</exception>
    /// <exception cref="OperationCanceledException">The cancellation token was cancelled.</exception>
    public bool Read() {
        if (++_tokenCount % CancellationInterval == 0) _cancellationToken.ThrowIfCancellationRequested();
        ReadOnlySpan<char> s = _text.Span;
        while (true) {
            switch (_state) {
                case State.LineStart:
                    if (SkipTrivia(s, newlines: true)) return true;
                    if (_pos >= s.Length) {
                        _state = State.End;
                        _tokenType = TomlTokenType.None;
                        return false;
                    }
                    if (s[_pos] == '[') {
                        _headerIsArray = _pos + 1 < s.Length && s[_pos + 1] == '[';
                        Token(_headerIsArray ? TomlTokenType.ArrayOfTablesHeaderStart : TomlTokenType.TableHeaderStart, _pos, _headerIsArray ? 2 : 1);
                        _state = State.HeaderKey;
                        return true;
                    }
                    ScanKey(s);
                    _state = State.AfterKey;
                    return true;
                case State.HeaderKey:
                    SkipSpaces(s);
                    ScanKey(s);
                    _state = State.HeaderAfterKey;
                    return true;
                case State.HeaderAfterKey:
                    SkipSpaces(s);
                    if (_pos < s.Length && s[_pos] == '.') {
                        _pos++;
                        _state = State.HeaderKey;
                        continue;
                    }
                    if (_pos < s.Length && s[_pos] == ']') {
                        if (_headerIsArray) {
                            if (_pos + 1 >= s.Length || s[_pos + 1] != ']') throw Error(TomlErrorCodes.UnexpectedToken, "An array-of-tables header must end with ']]'.", _pos);
                            Token(TomlTokenType.ArrayOfTablesHeaderEnd, _pos, 2);
                        } else {
                            Token(TomlTokenType.TableHeaderEnd, _pos, 1);
                        }
                        _state = State.AfterValue;
                        return true;
                    }
                    throw Error(TomlErrorCodes.UnexpectedToken, "Expected '.' or ']' in the table header.", _pos);
                case State.Key:
                    SkipSpaces(s);
                    ScanKey(s);
                    _state = State.AfterKey;
                    return true;
                case State.AfterKey:
                    SkipSpaces(s);
                    if (_pos < s.Length && s[_pos] == '.') {
                        _pos++;
                        _state = State.Key;
                        continue;
                    }
                    if (_pos < s.Length && s[_pos] == '=') {
                        _pos++;
                        _state = State.Value;
                        continue;
                    }
                    throw Error(TomlErrorCodes.UnexpectedToken, "Expected '=' or '.' after the key.", _pos);
                case State.Value:
                    SkipSpaces(s);
                    ScanValue(s);
                    return true;
                case State.AfterValue:
                    if (_depth == 0) {
                        if (SkipTrivia(s, newlines: false)) return true;
                        if (_pos < s.Length && s[_pos] != '\n') {
                            throw Error(TomlErrorCodes.UnexpectedToken, "Expected a newline or the end of the input after the value.", _pos);
                        }
                        _state = State.LineStart;
                        continue;
                    }
                    if (SkipTrivia(s, newlines: true)) return true;
                    if (_pos >= s.Length) throw Error(TomlErrorCodes.UnexpectedToken, Context == ArrayContext ? "The array is not closed." : "The inline table is not closed.", _pos);
                    if (s[_pos] == ',') {
                        _pos++;
                        _state = Context == ArrayContext ? State.ArrayItem : State.InlineKey;
                        continue;
                    }
                    if (Close(s)) return true;
                    throw Error(TomlErrorCodes.UnexpectedToken, Context == ArrayContext ? "Expected ',' or ']' in the array." : "Expected ',' or '}' in the inline table.", _pos);
                case State.ArrayFirst:
                case State.ArrayItem:
                    if (SkipTrivia(s, newlines: true)) return true;
                    if (_pos >= s.Length) throw Error(TomlErrorCodes.UnexpectedToken, "The array is not closed.", _pos);
                    if (Close(s)) return true;
                    ScanValue(s);
                    return true;
                case State.InlineFirst:
                case State.InlineKey:
                    if (SkipTrivia(s, newlines: true)) return true;
                    if (_pos >= s.Length) throw Error(TomlErrorCodes.UnexpectedToken, "The inline table is not closed.", _pos);
                    if (Close(s)) return true;
                    ScanKey(s);
                    _state = State.AfterKey;
                    return true;
                default:
                    _tokenType = TomlTokenType.None;
                    return false;
            }
        }
    }

    private byte Context => _contexts[_depth - 1];

    /// <summary>Emits the closing bracket of the innermost array or inline table when it is next.</summary>
    private bool Close(ReadOnlySpan<char> s) {
        char ch = s[_pos];
        if (Context == ArrayContext ? ch != ']' : ch != '}') return false;
        Token(ch == ']' ? TomlTokenType.EndArray : TomlTokenType.EndInlineTable, _pos, 1);
        _depth--;
        _state = State.AfterValue;
        return true;
    }

    private void Push(byte context) {
        if (_depth >= _options.MaxDepth) throw new TomlLimitExceededException(nameof(TomlReaderOptions.MaxDepth), _options.MaxDepth, Position(_pos));
        if (_depth == _contexts.Length) Array.Resize(ref _contexts, _contexts.Length * 2);
        _contexts[_depth++] = context;
    }

    private void Token(TomlTokenType type, int start, int length) {
        _tokenType = type;
        _tokenStart = start;
        _startLine = _line;
        _startColumn = start - _lineStart + 1;
        _pos = start + length;
        _tokenEnd = _pos;
        _endLine = _line;
        _endColumn = _pos - _lineStart + 1;
    }

    /// <summary>Records a token that may span lines; the line counters already point past it.</summary>
    private void Token(TomlTokenType type, int start, int line, int lineStart) {
        _tokenType = type;
        _tokenStart = start;
        _startLine = line;
        _startColumn = start - lineStart + 1;
        _tokenEnd = _pos;
        _endLine = _line;
        _endColumn = _pos - _lineStart + 1;
    }

    private void SkipSpaces(ReadOnlySpan<char> s) {
        while (_pos < s.Length && (s[_pos] == ' ' || s[_pos] == '\t')) _pos++;
    }

    /// <summary>
    /// Skips spaces, tabs, comments and (when allowed) newlines. Returns <see langword="true"/> when it stopped on a
    /// comment that the caller asked to see. Without newlines it stops at the first line break, leaving it unread.
    /// </summary>
    private bool SkipTrivia(ReadOnlySpan<char> s, bool newlines) {
        while (_pos < s.Length) {
            char ch = s[_pos];
            if (ch == ' ' || ch == '\t') {
                _pos++;
            } else if (ch == '\n') {
                if (!newlines) return false;
                _pos++;
                NewLine();
            } else if (ch == '\r') {
                if (_pos + 1 >= s.Length || s[_pos + 1] != '\n') throw Error(TomlErrorCodes.InvalidCharacter, "A carriage return must be followed by a line feed.", _pos);
                if (!newlines) {
                    _pos++;
                    return false;
                }
                _pos += 2;
                NewLine();
            } else if (ch == '#') {
                int start = _pos++;
                while (_pos < s.Length) {
                    char c = s[_pos];
                    if (c == '\n' || (c == '\r' && _pos + 1 < s.Length && s[_pos + 1] == '\n')) break;
                    if (TomlScalars.IsForbiddenControl(c)) throw Error(TomlErrorCodes.InvalidCharacter, "Control characters are not allowed in comments.", _pos);
                    _pos++;
                }
                if (_options.IncludeComments) {
                    Token(TomlTokenType.Comment, start, _pos - start);
                    return true;
                }
            } else {
                return false;
            }
        }
        return false;
    }

    private void NewLine() {
        _line++;
        _lineStart = _pos;
    }

    private void ScanKey(ReadOnlySpan<char> s) {
        int start = _pos;
        if (start >= s.Length) throw Error(TomlErrorCodes.InvalidKey, "Expected a key.", start);
        char ch = s[start];
        if (ch == '"' || ch == '\'') {
            if (start + 2 < s.Length && s[start + 1] == ch && s[start + 2] == ch) {
                throw Error(TomlErrorCodes.InvalidKey, "A multi-line string cannot be a key.", start);
            }
            if (ch == '"') ScanBasic(s, start);
            else ScanLiteral(s, start);
            _tokenType = TomlTokenType.Key;
            return;
        }
        int end = start;
        while (end < s.Length && TomlScalars.IsBareKeyChar(s[end])) end++;
        if (end == start) throw Error(TomlErrorCodes.InvalidKey, "Expected a key.", start);
        _style = TomlStringStyle.Bare;
        Token(TomlTokenType.Key, start, end - start);
    }

    private void ScanValue(ReadOnlySpan<char> s) {
        int start = _pos;
        if (start >= s.Length) throw Error(TomlErrorCodes.InvalidValue, "Expected a value.", start);
        char ch = s[start];
        switch (ch) {
            case '"':
                if (start + 2 < s.Length && s[start + 1] == '"' && s[start + 2] == '"') ScanMultilineBasic(s, start);
                else ScanBasic(s, start);
                _state = State.AfterValue;
                return;
            case '\'':
                if (start + 2 < s.Length && s[start + 1] == '\'' && s[start + 2] == '\'') ScanMultilineLiteral(s, start);
                else ScanLiteral(s, start);
                _state = State.AfterValue;
                return;
            case '[':
                Push(ArrayContext);
                Token(TomlTokenType.StartArray, start, 1);
                _state = State.ArrayFirst;
                return;
            case '{':
                Push(InlineContext);
                Token(TomlTokenType.StartInlineTable, start, 1);
                _state = State.InlineFirst;
                return;
        }
        int end = start;
        while (end < s.Length && TomlScalars.IsAtomChar(s[end])) end++;
        if (end == start) throw Error(TomlErrorCodes.InvalidValue, "Expected a value.", start);
        // A date followed by a space and a time is one date-time.
        if (end - start == 10 && s[start + 4] == '-' && end + 3 < s.Length && s[end] == ' '
            && TomlScalars.IsDigit(s[end + 1]) && TomlScalars.IsDigit(s[end + 2]) && s[end + 3] == ':') {
            end++;
            while (end < s.Length && TomlScalars.IsAtomChar(s[end])) end++;
        }
        TomlTokenType type = TomlScalars.Classify(s.Slice(start, end - start), out string? error, out int errorIndex);
        if (type == TomlTokenType.None) {
            string message = error switch {
                TomlErrorCodes.InvalidInteger => "The integer is malformed or outside the 64-bit range.",
                TomlErrorCodes.InvalidFloat => "The float is malformed.",
                TomlErrorCodes.InvalidDateTime => "The date or time is malformed or out of range.",
                _ => "Expected a value."
            };
            throw Error(error ?? TomlErrorCodes.InvalidValue, message, start + errorIndex);
        }
        Token(type, start, end - start);
        _state = State.AfterValue;
    }

    internal TomlPosition Position(int offset) {
        // Offsets are on the current line except for errors detected after the reader advanced past newlines.
        return offset >= _lineStart ? new TomlPosition(offset, _line, offset - _lineStart + 1) : PositionSlow(offset);
    }

    private TomlPosition PositionSlow(int offset) {
        ReadOnlySpan<char> s = _text.Span;
        int line = 1, lineStart = 0;
        for (int i = 0; i < offset && i < s.Length; i++) {
            if (s[i] == '\n') {
                line++;
                lineStart = i + 1;
            }
        }
        return new TomlPosition(offset, line, offset - lineStart + 1);
    }

    private TomlException Error(string code, string message, int offset) => new TomlException(code, message, Position(offset));
}
