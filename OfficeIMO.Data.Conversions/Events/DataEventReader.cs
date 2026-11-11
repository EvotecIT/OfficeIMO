namespace OfficeIMO.Data.Conversions;

/// <summary>
/// A forward-only pull reader over one source format, in the style of <see cref="System.Text.Json.Utf8JsonReader"/>.
/// <see cref="Text"/> is valid until the next <see cref="Read"/>. A reader over UTF-8 input can also hand scalars over as
/// UTF-8 (<see cref="Utf8Text"/>), in which case their characters are decoded only when <see cref="Text"/> is read.
/// </summary>
internal abstract class DataEventReader : IDisposable {
    private char[] _text = new char[256];
    private char[] _view;
    private int _viewStart;
    private int _textLength;
    private DuplicateKeyTracker? _keys;
    private byte[]? _utf8;
    private int _utf8Start;
    private int _utf8Length;
    private bool _decodePending;

    protected DataEventReader(DataConversionContext context) {
        Context = context;
        _view = _text;
    }

    public DataConversionContext Context { get; }

    public DataTokenType TokenType { get; private set; }

    public DataScalarKind ScalarKind { get; private set; }

    /// <summary>The property name, scalar text or comment of the current event.</summary>
    public ReadOnlySpan<char> Text {
        get {
            if (_decodePending) DecodeUtf8();
            return new ReadOnlySpan<char>(_view, _viewStart, _textLength);
        }
    }

    /// <summary>Whether the current scalar is available as UTF-8 through <see cref="Utf8Text"/>.</summary>
    public bool HasUtf8Text => _utf8 != null;

    /// <summary>The current scalar as UTF-8 when <see cref="HasUtf8Text"/> is set; valid until the next <see cref="Read"/>.</summary>
    public ReadOnlySpan<byte> Utf8Text => _utf8 == null ? default : new ReadOnlySpan<byte>(_utf8, _utf8Start, _utf8Length);

    /// <summary>
    /// Asks the reader to hand scalars over as UTF-8 where its input allows. Call before the first <see cref="Read"/>;
    /// returns whether the reader does.
    /// </summary>
    public virtual bool PreferUtf8() => false;

    /// <summary>A human-readable source position for diagnostics, such as <c>line 12, column 5</c>.</summary>
    public abstract string? Location { get; }

    /// <summary>Whether duplicate property names in one object are rejected. Sources whose format forbids them enable it.</summary>
    protected bool TrackDuplicateKeys {
        set => _keys = value ? new DuplicateKeyTracker() : null;
    }

    /// <summary>Advances to the next event; returns false at the end of the input.</summary>
    public bool Read() {
        Context.CancellationToken.ThrowIfCancellationRequested();
        if (!ReadCore()) {
            TokenType = DataTokenType.None;
            return false;
        }

        return true;
    }

    protected abstract bool ReadCore();

    protected void Set(DataTokenType type) {
        TokenType = type;
        _utf8 = null;
        _decodePending = false;
        _view = _text;
        _viewStart = 0;
        _textLength = 0;
        if (_keys != null) {
            if (type == DataTokenType.StartObject) _keys.Push();
            else if (type == DataTokenType.EndObject) _keys.Pop();
        }
    }

    protected void SetScalar(DataScalarKind kind, ReadOnlySpan<char> text) {
        TokenType = DataTokenType.Scalar;
        ScalarKind = kind;
        CopyText(text);
    }

    protected void SetScalar(DataScalarKind kind, string text) => SetScalar(kind, text.AsSpan());

    protected void SetPropertyName(ReadOnlySpan<char> name) {
        TokenType = DataTokenType.PropertyName;
        CopyText(name);
        if (_keys != null && !_keys.Add(name)) {
            throw new InvalidDataException("The key '" + name.ToString() + "' appears twice in one object" + At() + ".");
        }
    }

    protected void SetComment(ReadOnlySpan<char> text) {
        TokenType = DataTokenType.Comment;
        CopyText(text);
    }

    /// <summary>Gives a derived reader a buffer to decode the next text into; commit it with <see cref="CommitScalar(DataScalarKind, int)"/> or <see cref="CommitPropertyName(int)"/>.</summary>
    protected Span<char> GetTextBuffer(int minimumLength) {
        if (_text.Length < minimumLength) _text = new char[Math.Max(minimumLength, _text.Length * 2)];
        _view = _text;
        return _text;
    }

    protected void CommitScalar(DataScalarKind kind, int length) {
        _utf8 = null;
        _decodePending = false;
        TokenType = DataTokenType.Scalar;
        ScalarKind = kind;
        _view = _text;
        _viewStart = 0;
        _textLength = length;
    }

    /// <summary>Points the current scalar at text the derived reader already holds, without copying it.</summary>
    protected void CommitScalar(DataScalarKind kind, char[] buffer, int start, int length) {
        _utf8 = null;
        _decodePending = false;
        TokenType = DataTokenType.Scalar;
        ScalarKind = kind;
        _view = buffer;
        _viewStart = start;
        _textLength = length;
    }

    /// <summary>
    /// Points the current scalar at valid UTF-8 the derived reader already holds; its characters are decoded only if
    /// <see cref="Text"/> is read.
    /// </summary>
    protected void CommitUtf8Scalar(DataScalarKind kind, byte[] buffer, int start, int length) {
        TokenType = DataTokenType.Scalar;
        ScalarKind = kind;
        _utf8 = buffer;
        _utf8Start = start;
        _utf8Length = length;
        _decodePending = true;
    }

    /// <summary>Points the current property name at text the derived reader already holds, without copying it.</summary>
    protected void CommitPropertyName(char[] buffer, int start, int length) {
        _utf8 = null;
        _decodePending = false;
        TokenType = DataTokenType.PropertyName;
        _view = buffer;
        _viewStart = start;
        _textLength = length;
        if (_keys != null && !_keys.Add(Text)) {
            throw new InvalidDataException("The key '" + Text.ToString() + "' appears twice in one object" + At() + ".");
        }
    }

    protected void CommitPropertyName(int length) {
        _utf8 = null;
        _decodePending = false;
        TokenType = DataTokenType.PropertyName;
        _view = _text;
        _viewStart = 0;
        _textLength = length;
        if (_keys != null && !_keys.Add(Text)) {
            throw new InvalidDataException("The key '" + Text.ToString() + "' appears twice in one object" + At() + ".");
        }
    }

    /// <summary>Whether the innermost object already has this key.</summary>
    protected bool HasKey(ReadOnlySpan<char> key) => _keys != null && _keys.Contains(key);

    protected string At() => Location is { } location ? " at " + location : string.Empty;

    private void DecodeUtf8() {
        _decodePending = false;
        int count = Encoding.UTF8.GetMaxCharCount(_utf8Length);
        if (_text.Length < count) _text = new char[Math.Max(count, _text.Length * 2)];
        _view = _text;
        _viewStart = 0;
        _textLength = Encoding.UTF8.GetChars(_utf8!, _utf8Start, _utf8Length, _text, 0);
    }

    private void CopyText(ReadOnlySpan<char> text) {
        _utf8 = null;
        _decodePending = false;
        if (_text.Length < text.Length) _text = new char[Math.Max(text.Length, _text.Length * 2)];
        text.CopyTo(_text);
        _view = _text;
        _viewStart = 0;
        _textLength = text.Length;
    }

    public virtual void Dispose() {
    }
}
