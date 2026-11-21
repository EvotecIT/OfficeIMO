namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Reads JSON or JSON Lines from a stream with <see cref="Utf8JsonReader"/> over one pooled, refilled buffer, so memory
/// stays bounded by the longest token. Tokens are decoded in batches into a reused character arena, which amortizes
/// creating the reader and lets property names and strings reach the sink without a copy. Each top-level value is a
/// document. On .NET 8 and later a consumer that writes UTF-8 can take scalars as UTF-8 (<see cref="DataEventReader.PreferUtf8"/>): they
/// point into the input buffer, or into a byte arena when escaped, and are decoded to characters only on request.
/// </summary>
internal sealed class JsonEventReader : DataEventReader {
    private const int InitialBufferSize = 64 * 1024;
    private const int BatchTokens = 4096;
    private readonly Stream _stream;
    private readonly bool _leaveOpen;
    private readonly bool _jsonLines;
    private readonly JsonReaderOptions _options;
    private byte[] _buffer;
    private int _start;
    private int _end;
    private bool _final;
    private bool _bomChecked;
    private long _offset;
    private JsonReaderState _state;
    private bool _atBoundary = true;
    private int _scanDepth;

    // The current batch: token kinds and their text ranges in the arena.
    private Token[] _tokens = new Token[BatchTokens];
    private int _tokenCount;
    private int _tokenIndex;
    private char[] _arena = new char[16 * 1024];
    private int _arenaLength;
    private long _tokenOffset;
#if NET8_0_OR_GREATER
    private static readonly byte[] TrueUtf8 = "true"u8.ToArray();
    private static readonly byte[] FalseUtf8 = "false"u8.ToArray();
    private bool _utf8;
    private byte[] _bytes = Array.Empty<byte>();
    private int _bytesLength;
#endif

    // Document framing.
    private int _depth;
    private bool _inDocument;
    private bool _endDocumentPending;
    private bool _sawDocument;

    // Where a token's text lives: characters in the arena, or UTF-8 in the input buffer, the byte arena or a literal.
    private enum TextSource : byte {
        Chars,
        Input,
        Bytes,
        True,
        False
    }

    private struct Token {
        public DataTokenType Type;
        public DataScalarKind Kind;
        public TextSource Source;
        public int Start;
        public int Length;
        public long Offset;
    }

    public JsonEventReader(Stream stream, bool jsonLines, DataConversionContext context, bool leaveOpen = true) : base(context) {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
        _jsonLines = jsonLines;
        _buffer = ArrayPool<byte>.Shared.Rent(InitialBufferSize);
        _options = context.Options.JsonReader ?? new JsonReaderOptions { MaxDepth = 256 };
        _state = new JsonReaderState(_options);
        TrackDuplicateKeys = true;
    }

#if NET8_0_OR_GREATER
    public override bool PreferUtf8() => _utf8 = true;
#endif

    public override string? Location => "byte " + _tokenOffset.ToString(CultureInfo.InvariantCulture);

    protected override bool ReadCore() {
        if (_endDocumentPending) {
            _endDocumentPending = false;
            _inDocument = false;
            Set(DataTokenType.EndDocument);
            return true;
        }

        if (_tokenIndex == _tokenCount && !FillBatch()) {
            if (!_sawDocument && !_jsonLines) throw new InvalidDataException("The JSON input is empty.");
            return false;
        }

        ref Token token = ref _tokens[_tokenIndex];
        _tokenOffset = token.Offset;
        if (!_inDocument && token.Type != DataTokenType.Comment) {
            _inDocument = true;
            _sawDocument = true;
            Set(DataTokenType.StartDocument);
            return true;
        }

        _tokenIndex++;
        switch (token.Type) {
            case DataTokenType.StartObject:
            case DataTokenType.StartArray:
                _depth++;
                Set(token.Type);
                break;
            case DataTokenType.EndObject:
            case DataTokenType.EndArray:
                _depth--;
                Set(token.Type);
                if (_depth == 0) _endDocumentPending = true;
                break;
            case DataTokenType.PropertyName:
                CommitPropertyName(_arena, token.Start, token.Length);
                break;
            case DataTokenType.Scalar:
                switch (token.Source) {
#if NET8_0_OR_GREATER
                    case TextSource.Input: CommitUtf8Scalar(token.Kind, _buffer, token.Start, token.Length); break;
                    case TextSource.Bytes: CommitUtf8Scalar(token.Kind, _bytes, token.Start, token.Length); break;
                    case TextSource.True: CommitUtf8Scalar(token.Kind, TrueUtf8, 0, 4); break;
                    case TextSource.False: CommitUtf8Scalar(token.Kind, FalseUtf8, 0, 5); break;
#endif
                    default: CommitScalar(token.Kind, _arena, token.Start, token.Length); break;
                }

                if (_depth == 0) _endDocumentPending = true;
                break;
            case DataTokenType.Comment:
                SetComment(new ReadOnlySpan<char>(_arena, token.Start, token.Length));
                break;
        }

        return true;
    }

    // Decodes the next run of tokens with one reader instance. JSON Lines batches stop at the end of each top-level
    // value, because the next value starts with a fresh reader state.
    private bool FillBatch() {
        _tokenCount = 0;
        _tokenIndex = 0;
        _arenaLength = 0;
#if NET8_0_OR_GREATER
        _bytesLength = 0;
#endif
        while (true) {
            if (_jsonLines && _atBoundary) {
                if (!SkipToNextValue()) return false;
                _atBoundary = false;
                _state = new JsonReaderState(_options);
            }

            ReadTokens();
            if (_tokenCount > 0) return true;
            if (_final) return false;
            Refill();
        }
    }

    private void ReadTokens() {
        if (_end - _start == 0 && !_final) return;
        var reader = new Utf8JsonReader(new ReadOnlySpan<byte>(_buffer, _start, _end - _start), _final, _state);
        try {
            while (_tokenCount < BatchTokens && _arenaLength < _arena.Length - 16 && reader.Read()) {
                Add(ref reader);
                if (_jsonLines && _scanDepth == 0 && reader.TokenType is not (JsonTokenType.PropertyName or JsonTokenType.Comment)) {
                    _atBoundary = true;
                    break;
                }
            }
        } catch (JsonException error) {
            throw new InvalidDataException("The JSON input is invalid near byte "
                + (_offset + _start + reader.BytesConsumed).ToString(CultureInfo.InvariantCulture) + ": " + error.Message, error);
        } catch (InvalidOperationException error) {
            throw new InvalidDataException("The JSON input has invalid text near byte "
                + (_offset + _start + reader.BytesConsumed).ToString(CultureInfo.InvariantCulture) + ": " + error.Message, error);
        }

        _state = reader.CurrentState;
        _start += (int)reader.BytesConsumed;
    }

    private void Add(ref Utf8JsonReader reader) {
        ref Token token = ref _tokens[_tokenCount++];
        token.Offset = _offset + _start + reader.TokenStartIndex;
        token.Kind = default;
        token.Source = TextSource.Chars;
        token.Start = _arenaLength;
        token.Length = 0;
#if NET8_0_OR_GREATER
        if (_utf8 && AddUtf8(ref reader, ref token)) return;
#endif
        switch (reader.TokenType) {
            case JsonTokenType.StartObject: token.Type = DataTokenType.StartObject; _scanDepth++; return;
            case JsonTokenType.StartArray: token.Type = DataTokenType.StartArray; _scanDepth++; return;
            case JsonTokenType.EndObject: token.Type = DataTokenType.EndObject; _scanDepth--; return;
            case JsonTokenType.EndArray: token.Type = DataTokenType.EndArray; _scanDepth--; return;
            case JsonTokenType.PropertyName:
                token.Type = DataTokenType.PropertyName;
                token.Length = reader.CopyString(Arena(reader.ValueSpan.Length));
                break;
            case JsonTokenType.String:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.String;
                token.Length = reader.CopyString(Arena(reader.ValueSpan.Length));
                break;
            case JsonTokenType.Number: {
                token.Type = DataTokenType.Scalar;
                ReadOnlySpan<byte> raw = reader.ValueSpan;
                Span<char> text = Arena(raw.Length);
                bool integer = true;
                for (int i = 0; i < raw.Length; i++) {
                    byte b = raw[i];
                    if (b == (byte)'.' || b == (byte)'e' || b == (byte)'E') integer = false;
                    text[i] = (char)b;
                }

                token.Kind = integer ? DataScalarKind.Integer : DataScalarKind.Float;
                token.Length = raw.Length;
                break;
            }
            case JsonTokenType.True:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.Boolean;
                token.Length = Literal("true");
                break;
            case JsonTokenType.False:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.Boolean;
                token.Length = Literal("false");
                break;
            case JsonTokenType.Null:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.Null;
                break;
            case JsonTokenType.Comment:
                token.Type = DataTokenType.Comment;
                token.Length = Literal(reader.GetComment());
                break;
            default:
                throw new InvalidDataException("Unexpected JSON token " + reader.TokenType + ".");
        }

        _arenaLength += token.Length;
    }

#if NET8_0_OR_GREATER
    // Scalars as UTF-8: unescaped strings and numbers stay in the input buffer, which is not moved until the batch is
    // consumed; escaped strings are unescaped into the byte arena. Strings are validated here, because only decoding
    // them would.
    private bool AddUtf8(ref Utf8JsonReader reader, ref Token token) {
        switch (reader.TokenType) {
            case JsonTokenType.String:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.String;
                if (reader.ValueIsEscaped) {
                    int room = reader.ValueSpan.Length;
                    if (_bytes.Length - _bytesLength < room) Array.Resize(ref _bytes, Math.Max(Math.Max(_bytes.Length * 2, 4096), _bytesLength + room));
                    int written = reader.CopyString(new Span<byte>(_bytes, _bytesLength, room));
                    SetUtf8(ref token, TextSource.Bytes, _bytesLength, written);
                    _bytesLength += written;
                } else {
                    SetUtf8(ref token, TextSource.Input, InputOffset(reader.ValueSpan), reader.ValueSpan.Length);
                }

                return true;
            case JsonTokenType.Number: {
                token.Type = DataTokenType.Scalar;
                ReadOnlySpan<byte> raw = reader.ValueSpan;
                token.Kind = raw.IndexOfAny((byte)'.', (byte)'e', (byte)'E') < 0 ? DataScalarKind.Integer : DataScalarKind.Float;
                token.Source = TextSource.Input;
                token.Start = InputOffset(raw);
                token.Length = raw.Length;
                return true;
            }
            case JsonTokenType.True:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.Boolean;
                token.Source = TextSource.True;
                return true;
            case JsonTokenType.False:
                token.Type = DataTokenType.Scalar;
                token.Kind = DataScalarKind.Boolean;
                token.Source = TextSource.False;
                return true;
            default:
                return false;
        }
    }

    private void SetUtf8(ref Token token, TextSource source, int start, int length) {
        if (!System.Text.Unicode.Utf8.IsValid(new ReadOnlySpan<byte>(source == TextSource.Input ? _buffer : _bytes, start, length))) {
            throw new InvalidDataException("The JSON input has invalid text near byte " + token.Offset.ToString(CultureInfo.InvariantCulture)
                + ": a string is not valid UTF-8.");
        }

        token.Source = source;
        token.Start = start;
        token.Length = length;
    }

    private int InputOffset(ReadOnlySpan<byte> value) =>
        (int)System.Runtime.CompilerServices.Unsafe.ByteOffset(ref _buffer[0], ref System.Runtime.InteropServices.MemoryMarshal.GetReference(value));
#endif

    // Room for a value's characters at the end of the arena; decoded text is never longer than its UTF-8 bytes.
    private Span<char> Arena(int length) {
        if (_arena.Length - _arenaLength < length) Array.Resize(ref _arena, Math.Max(_arena.Length * 2, _arenaLength + length));
        return new Span<char>(_arena, _arenaLength, length);
    }

    private int Literal(string value) {
        value.AsSpan().CopyTo(Arena(value.Length));
        return value.Length;
    }

    // Skips whitespace between JSON Lines values; false when only whitespace remains.
    private bool SkipToNextValue() {
        while (true) {
            while (_start < _end) {
                byte b = _buffer[_start];
                if (b != (byte)' ' && b != (byte)'\t' && b != (byte)'\n' && b != (byte)'\r') return true;
                _start++;
            }

            if (_final) return false;
            Refill();
        }
    }

    private void Refill() {
        int remaining = _end - _start;
        if (_start > 0) {
            Buffer.BlockCopy(_buffer, _start, _buffer, 0, remaining);
            _offset += _start;
            _start = 0;
            _end = remaining;
        }

        if (_end == _buffer.Length) {
            byte[] larger = ArrayPool<byte>.Shared.Rent(_buffer.Length * 2);
            Buffer.BlockCopy(_buffer, 0, larger, 0, _end);
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = larger;
        }

        int read = _stream.Read(_buffer, _end, _buffer.Length - _end);
        if (read == 0) _final = true;
        else _end += read;
        if (!_bomChecked && _end >= 3) {
            _bomChecked = true;
            if (_offset == 0 && _start == 0 && _buffer[0] == 0xEF && _buffer[1] == 0xBB && _buffer[2] == 0xBF) _start = 3;
        } else if (_final) {
            _bomChecked = true;
        }
    }

    public override void Dispose() {
        if (_buffer.Length > 0) {
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = Array.Empty<byte>();
        }

        if (!_leaveOpen) _stream.Dispose();
    }
}
