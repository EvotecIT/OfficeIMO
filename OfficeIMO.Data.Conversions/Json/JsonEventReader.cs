namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Reads JSON or JSON Lines from a stream with <see cref="Utf8JsonReader"/> over one pooled, refilled buffer, so memory
/// stays bounded by the longest token. Each top-level value is a document.
/// </summary>
internal sealed class JsonEventReader : DataEventReader {
    private const int InitialBufferSize = 64 * 1024;
    private readonly Stream _stream;
    private readonly bool _leaveOpen;
    private readonly bool _jsonLines;
    private byte[] _buffer;
    private int _start;
    private int _end;
    private bool _final;
    private long _offset;
    private long _tokenOffset;
    private JsonReaderState _state;
    private int _depth;
    private bool _inDocument;
    private bool _endDocumentPending;
    private bool _held;
    private DataTokenType _heldType;
    private DataScalarKind _heldKind;
    private int _heldLength;
    private bool _sawDocument;
    private readonly JsonReaderOptions _options;
    private bool _atBoundary = true;

    public JsonEventReader(Stream stream, bool jsonLines, DataConversionContext context, bool leaveOpen = true) : base(context) {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
        _jsonLines = jsonLines;
        _buffer = ArrayPool<byte>.Shared.Rent(InitialBufferSize);
        _options = context.Options.JsonReader ?? new JsonReaderOptions { MaxDepth = 256 };
        _state = new JsonReaderState(_options);
        TrackDuplicateKeys = true;
    }

    public override string? Location => "byte " + _tokenOffset.ToString(CultureInfo.InvariantCulture);

    protected override bool ReadCore() {
        if (_endDocumentPending) {
            _endDocumentPending = false;
            _inDocument = false;
            Set(DataTokenType.EndDocument);
            return true;
        }

        if (_held) {
            _held = false;
            Deliver(_heldType, _heldKind, _heldLength);
            return true;
        }

        if (!ReadToken(out DataTokenType type, out DataScalarKind kind, out int length)) {
            if (!_sawDocument && !_jsonLines) throw new InvalidDataException("The JSON input is empty.");
            return false;
        }

        if (!_inDocument && type != DataTokenType.Comment) {
            _inDocument = true;
            _sawDocument = true;
            _held = true;
            _heldType = type;
            _heldKind = kind;
            _heldLength = length;
            Set(DataTokenType.StartDocument);
            return true;
        }

        Deliver(type, kind, length);
        return true;
    }

    private void Deliver(DataTokenType type, DataScalarKind kind, int length) {
        switch (type) {
            case DataTokenType.StartObject:
            case DataTokenType.StartArray:
                _depth++;
                Set(type);
                break;
            case DataTokenType.EndObject:
            case DataTokenType.EndArray:
                _depth--;
                Set(type);
                if (_depth == 0) EndValue();
                break;
            case DataTokenType.PropertyName:
                CommitPropertyName(length);
                break;
            case DataTokenType.Scalar:
                CommitScalar(kind, length);
                if (_depth == 0) EndValue();
                break;
            case DataTokenType.Comment:
                SetComment(GetTextBuffer(length).Slice(0, length));
                break;
        }
    }

    // JSON Lines: each top-level value starts a fresh reader state, so values can follow each other on new lines.
    private void EndValue() {
        _endDocumentPending = true;
        if (_jsonLines) {
            _state = new JsonReaderState(_options);
            _atBoundary = true;
        }
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

    private bool ReadToken(out DataTokenType type, out DataScalarKind kind, out int length) {
        if (_jsonLines && _atBoundary) {
            _atBoundary = false;
            if (!SkipToNextValue()) {
                type = DataTokenType.None;
                kind = default;
                length = 0;
                return false;
            }
        }

        while (true) {
            var reader = new Utf8JsonReader(new ReadOnlySpan<byte>(_buffer, _start, _end - _start), _final, _state);
            bool read;
            try {
                read = reader.Read();
            } catch (JsonException error) {
                throw new InvalidDataException("The JSON input is invalid near byte "
                    + (_offset + _start + reader.BytesConsumed).ToString(CultureInfo.InvariantCulture) + ": " + error.Message, error);
            }

            if (read) {
                _tokenOffset = _offset + _start + reader.TokenStartIndex;
                try {
                    Extract(ref reader, out type, out kind, out length);
                } catch (InvalidOperationException error) {
                    throw new InvalidDataException("The JSON input has invalid text at byte " + _tokenOffset.ToString(CultureInfo.InvariantCulture) + ": " + error.Message, error);
                }

                _state = reader.CurrentState;
                _start += (int)reader.BytesConsumed;
                return true;
            }

            _state = reader.CurrentState;
            _start += (int)reader.BytesConsumed;
            if (_final) {
                type = DataTokenType.None;
                kind = default;
                length = 0;
                return false;
            }

            Refill();
        }
    }

    private void Extract(ref Utf8JsonReader reader, out DataTokenType type, out DataScalarKind kind, out int length) {
        kind = default;
        length = 0;
        switch (reader.TokenType) {
            case JsonTokenType.StartObject: type = DataTokenType.StartObject; return;
            case JsonTokenType.EndObject: type = DataTokenType.EndObject; return;
            case JsonTokenType.StartArray: type = DataTokenType.StartArray; return;
            case JsonTokenType.EndArray: type = DataTokenType.EndArray; return;
            case JsonTokenType.PropertyName:
                type = DataTokenType.PropertyName;
                length = reader.CopyString(GetTextBuffer(reader.ValueSpan.Length));
                return;
            case JsonTokenType.String:
                type = DataTokenType.Scalar;
                kind = DataScalarKind.String;
                length = reader.CopyString(GetTextBuffer(reader.ValueSpan.Length));
                return;
            case JsonTokenType.Number: {
                type = DataTokenType.Scalar;
                ReadOnlySpan<byte> raw = reader.ValueSpan;
                Span<char> text = GetTextBuffer(raw.Length);
                bool integer = true;
                for (int i = 0; i < raw.Length; i++) {
                    byte b = raw[i];
                    if (b == (byte)'.' || b == (byte)'e' || b == (byte)'E') integer = false;
                    text[i] = (char)b;
                }

                length = raw.Length;
                kind = integer ? DataScalarKind.Integer : DataScalarKind.Float;
                return;
            }
            case JsonTokenType.True:
                type = DataTokenType.Scalar;
                kind = DataScalarKind.Boolean;
                length = Write("true");
                return;
            case JsonTokenType.False:
                type = DataTokenType.Scalar;
                kind = DataScalarKind.Boolean;
                length = Write("false");
                return;
            case JsonTokenType.Null:
                type = DataTokenType.Scalar;
                kind = DataScalarKind.Null;
                return;
            case JsonTokenType.Comment: {
                type = DataTokenType.Comment;
                string comment = reader.GetComment();
                length = Write(comment);
                return;
            }
            default:
                throw new InvalidDataException("Unexpected JSON token " + reader.TokenType + ".");
        }
    }

    private int Write(string value) {
        value.AsSpan().CopyTo(GetTextBuffer(value.Length));
        return value.Length;
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
        if (_offset == 0 && _end >= 3 && _buffer[0] == 0xEF && _buffer[1] == 0xBB && _buffer[2] == 0xBF && _start == 0 && !_bomSkipped) {
            _bomSkipped = true;
            _start = 3;
        }
    }

    private bool _bomSkipped;

    public override void Dispose() {
        if (_buffer.Length > 0) {
            ArrayPool<byte>.Shared.Return(_buffer);
            _buffer = Array.Empty<byte>();
        }

        if (!_leaveOpen) _stream.Dispose();
    }
}
