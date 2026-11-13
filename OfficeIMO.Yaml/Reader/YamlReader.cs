using System.Runtime.InteropServices;

namespace OfficeIMO.Yaml;

/// <summary>
/// A forward-only, low-allocation pull reader for YAML 1.2 text, modeled on <c>Utf8JsonReader</c> and <c>XmlReader</c>.
/// Each <see cref="Read"/> advances to the next event of the serialization tree: stream, document, collection
/// start and end, scalar, alias and (optionally) comment.
/// </summary>
/// <remarks>
/// <para>Reading allocates nothing per event. Strings are created only when <see cref="GetString"/>, <see cref="Tag"/>
/// or <see cref="Anchor"/> is called, and plain scalars without folding are copied straight from the source.</para>
/// <para><see cref="ValueSpan"/> and the other span-based members are valid only until the next <see cref="Read"/>.
/// Positions are precise enough to rebuild the input: the text between consecutive events is whitespace, comments,
/// line breaks and the indicators <c>-</c>, <c>?</c>, <c>:</c> and <c>,</c>.</para>
/// <para>Malformed input throws <see cref="YamlException"/>; after an exception the reader cannot continue.
/// A reader is not thread-safe. Dispose it to return pooled buffers; the reader never disposes a caller's stream or
/// text reader.</para>
/// </remarks>
public sealed partial class YamlReader : IDisposable {
    private readonly YamlScanContext _c;
    private readonly ReadOnlyMemory<char> _memory;
    // A whole-string input, sliced directly: ReadOnlyMemory<char>.Span checks its backing object on every access.
    private readonly string? _string;
    private readonly YamlReaderOptions _options;
    private readonly YamlStreamSource? _streamSource;
    private YamlEvent _event;
    // The event held back while the comments before it are read; created only when comment events are requested.
    private YamlEvent[]? _pending;
    private bool _hasPending;
    private bool _finished;
    private bool _faulted;
    private bool _disposed;
    private string? _stringValue;
    private char[]? _scratch;
    private List<OfficeDiagnostic>? _diagnostics;

    /// <summary>Creates a reader over YAML text.</summary>
    /// <param name="yaml">The YAML text. A leading U+FEFF byte-order mark is skipped.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Checked at bounded intervals while scanning.</param>
    public YamlReader(string yaml, YamlReaderOptions? options = null, CancellationToken cancellationToken = default)
        : this((yaml ?? throw new ArgumentNullException(nameof(yaml))).AsMemory(), options, cancellationToken) {
    }

    /// <summary>Creates a reader over YAML text held in memory. The memory must not change while the reader is used.</summary>
    /// <param name="yaml">The YAML text. A leading U+FEFF byte-order mark is skipped.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Checked at bounded intervals while scanning.</param>
    public YamlReader(ReadOnlyMemory<char> yaml, YamlReaderOptions? options = null, CancellationToken cancellationToken = default) {
        _options = options ?? YamlReaderOptions.Default;
        _c = new YamlScanContext(_options, cancellationToken);
        _memory = yaml;
        if (MemoryMarshal.TryGetString(yaml, out string? text, out int start, out int length) && start == 0 && length == text.Length) _string = text;
        if (_options.MaxInputBytes is long max && yaml.Length > max) {
            _c.Release();
            throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxInputBytes), max, new YamlPosition(max, 0, 0));
        }
    }

    /// <summary>
    /// Creates a reader over a caller-owned <see cref="TextReader"/>. Text is read into a pooled sliding buffer that keeps
    /// only what the current event needs; it grows only while a single token is longer than the buffer.
    /// </summary>
    /// <param name="reader">The source of YAML text. It is not disposed by the reader.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Checked at bounded intervals while scanning and reading.</param>
    public YamlReader(TextReader reader, YamlReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (reader is null) throw new ArgumentNullException(nameof(reader));
        _options = options ?? YamlReaderOptions.Default;
        _c = new YamlScanContext(_options, cancellationToken);
        _c.AttachSource(new YamlTextReaderSource(reader), _options.MaxInputBytes ?? long.MaxValue);
    }

    /// <summary>
    /// Creates a reader over a caller-owned byte stream. The encoding (UTF-8, UTF-16 or UTF-32, either byte order) is
    /// detected from a byte-order mark or the first bytes as the YAML specification requires; invalid bytes throw
    /// <see cref="YamlException"/> with code <see cref="YamlErrorCodes.InvalidEncoding"/>. Decoded text is read into a
    /// pooled sliding buffer.
    /// </summary>
    /// <param name="stream">A readable stream. It is not disposed by the reader.</param>
    /// <param name="options">Reader settings, or <see langword="null"/> for the defaults.</param>
    /// <param name="cancellationToken">Checked at bounded intervals while scanning and reading.</param>
    public YamlReader(Stream stream, YamlReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream is null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("The stream must be readable.", nameof(stream));
        _options = options ?? YamlReaderOptions.Default;
        _c = new YamlScanContext(_options, cancellationToken);
        _streamSource = new YamlStreamSource(stream);
        _c.AttachSource(_streamSource, _options.MaxInputBytes ?? long.MaxValue);
    }

    /// <summary>Gets the type of the current event, or <see cref="YamlTokenType.None"/> before the first <see cref="Read"/> and after the last.</summary>
    public YamlTokenType TokenType => _event.Type;

    /// <summary>Gets the style of the current scalar or collection, or <see cref="YamlStyle.None"/> for other events.</summary>
    public YamlStyle Style => _event.Style;

    /// <summary>
    /// Gets whether the current document start or end has no marker (<c>---</c> or <c>...</c>) in the source, or whether
    /// the current mapping is an implicit single-pair mapping inside a flow sequence (<c>[ key: value ]</c>).
    /// </summary>
    public bool IsImplicit => _event.Implicit;

    /// <summary>Gets where the current event starts. Node events start at their first property (anchor or tag).</summary>
    public YamlPosition Start => new YamlPosition(_event.Start, _event.StartLine + 1, _event.StartColumn + 1);

    /// <summary>
    /// Gets where the current event ends. Block collection ends are zero-width at the end of their last content;
    /// empty scalars are zero-width where the omitted node would be.
    /// </summary>
    public YamlPosition End => new YamlPosition(_event.End, _event.EndLine + 1, _event.EndColumn + 1);

    /// <summary>
    /// Gets the raw source of the current value without decoding: a scalar's text between its quotes (escapes intact)
    /// or a block scalar's content lines (indentation intact), a comment's text after <c>#</c>, or an alias's name.
    /// Empty for other events. Valid until the next <see cref="Read"/>; use <see cref="GetString"/> for the value.
    /// </summary>
    public ReadOnlySpan<char> ValueSpan {
        get {
            switch (_event.Type) {
                case YamlTokenType.Scalar:
                case YamlTokenType.Comment:
                    return Slice(_event.ValueStart, _event.ValueEnd);
                case YamlTokenType.Alias:
                    return Slice(_event.AnchorStart, _event.AnchorEnd);
                default:
                    return default;
            }
        }
    }

    /// <summary>
    /// Gets the schema the current document resolves untagged plain scalars with: <see cref="YamlReaderOptions.Schema"/>
    /// when set; otherwise <see cref="YamlSchema.Yaml11"/> for a document that declares <c>%YAML 1.1</c> and
    /// <see cref="YamlSchema.Core"/> for any other. It changes at each <see cref="YamlTokenType.DocumentStart"/>, so a
    /// streaming consumer resolves values exactly as a loaded <see cref="YamlDocument"/> does.
    /// </summary>
    public YamlSchema Schema => _documentSchema;

    /// <summary>
    /// Gets the findings so far, in source order: each indentation the specification rejects that the reader accepted
    /// because <see cref="YamlReaderOptions.Strict"/> is off (<see cref="OfficeDiagnosticCodes.YamlLenientIndentation"/>),
    /// once per outermost flow collection and once per quoted scalar, with a line and column as the location. Empty, and
    /// free, while there are none; the list grows as reading continues.
    /// </summary>
    public IReadOnlyList<OfficeDiagnostic> Diagnostics {
        get {
            List<YamlLenientFinding>? findings = _c.Findings;
            if (findings is null) return Array.Empty<OfficeDiagnostic>();
            List<OfficeDiagnostic> diagnostics = _diagnostics ??= new List<OfficeDiagnostic>(findings.Count);
            for (int i = diagnostics.Count; i < findings.Count; i++) diagnostics.Add(findings[i].ToDiagnostic(null));
            return diagnostics;
        }
    }

    /// <summary>Gets the raw lenient-indentation findings so far, or <see langword="null"/>, for the document model.</summary>
    internal List<YamlLenientFinding>? LenientFindings => _c.Findings;

    /// <summary>
    /// Advances to the next event. Returns <see langword="false"/> after <see cref="YamlTokenType.StreamEnd"/> has been read.
    /// </summary>
    /// <exception cref="YamlException">The input is malformed or exceeds a limit.</exception>
    /// <exception cref="OperationCanceledException">The cancellation token was signaled.</exception>
    public bool Read() {
        if (_disposed) throw new ObjectDisposedException(nameof(YamlReader));
        if (_faulted) throw new InvalidOperationException("The reader cannot continue after an error.");
        _stringValue = null;
        if (_finished) {
            _event = default;
            return false;
        }
        try {
            ReadCore();
        } catch (YamlException exception) when (_streamSource is null && exception.AttachSource(_memory.ToString())) {
            throw;
        } catch (Exception) {
            _faulted = true;
            throw;
        }
        return true;
    }

    private void ReadCore() {
        YamlScanContext c = _c;
        c.RetainMark = -1;
        if (c.Source is not null) c.Compact();
        if (_hasPending) {
            if (c.CommentCount > 0 && c.Comments[c.CommentHead].Start < _pending![0].Start) {
                TakeComment();
                return;
            }
            _event = _pending![0];
            _hasPending = false;
        } else {
            var scanner = new YamlScanner(CurrentText, c);
            ParseEvent(ref scanner, out _event);
            if (c.CommentCount > 0 && c.Comments[c.CommentHead].Start <= _event.Start) {
                _pending ??= new YamlEvent[1];
                _pending[0] = _event;
                _hasPending = true;
                TakeComment();
                return;
            }
        }
        if (_event.Type == YamlTokenType.StreamEnd) _finished = true;
    }

    private void TakeComment() {
        YamlScanContext c = _c;
        YamlComment comment = c.Comments[c.CommentHead++];
        if (--c.CommentCount == 0) c.CommentHead = 0;
        _event = default;
        _event.Type = YamlTokenType.Comment;
        _event.Start = comment.Start;
        _event.End = comment.End;
        _event.StartLine = _event.EndLine = comment.Line;
        _event.StartColumn = comment.Column;
        _event.EndColumn = comment.Column + (comment.End - comment.Start);
        _event.ValueStart = comment.Start + 1;
        _event.ValueEnd = comment.End;
        _event.AnchorStart = _event.AnchorEnd = -1;
    }

    private ReadOnlySpan<char> CurrentText => _string is not null ? _string.AsSpan() : _c.Source is null ? _memory.Span : new ReadOnlySpan<char>(_c.Buffer, 0, _c.Length);

    private ReadOnlySpan<char> Slice(int start, int end) => start < 0 ? default : CurrentText.Slice(start - _c.Base, end - start);

    /// <summary>Gets the detected source encoding for stream input (UTF-8 for text input), for lossless saving.</summary>
    internal YamlTextEncoding SourceEncoding => _streamSource?.Encoding ?? YamlTextEncoding.Utf8;

    /// <summary>Gets whether stream input started with a byte-order mark, for lossless saving.</summary>
    internal bool SourceHasByteOrderMark => _streamSource?.HasByteOrderMark ?? false;

    /// <summary>Gets the schema in force for the current document.</summary>
    internal YamlSchema DocumentSchema => _documentSchema;

    /// <summary>Gets whether the current scalar is a mapping key.</summary>
    internal bool IsKey => _event.IsKey;

    /// <summary>Gets the current event for the document model, which builds nodes from its offsets.</summary>
    internal ref readonly YamlEvent CurrentEvent => ref _event;

    /// <summary>Gets the current document's <c>%YAML</c> version as <c>major.minor</c>, or <see langword="null"/> without the directive.</summary>
    internal string? DocumentVersion => _versionMajor < 0
        ? null
        : _versionMajor.ToString(CultureInfo.InvariantCulture) + "." + _versionMinor.ToString(CultureInfo.InvariantCulture);

    /// <summary>Copies the current document's <c>%TAG</c> directives, or returns an empty array.</summary>
    internal YamlTagDirective[] CopyTagDirectives() {
        if (_tagDirectiveCount == 0) return Array.Empty<YamlTagDirective>();
        var copy = new YamlTagDirective[_tagDirectiveCount];
        Array.Copy(_tagDirectives, copy, _tagDirectiveCount);
        return copy;
    }

    /// <summary>Returns pooled buffers. The caller's stream or text reader is not disposed.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _c.Release();
        _strings?.Release();
        _strings = null;
        if (_scratch is not null) {
            ArrayPool<char>.Shared.Return(_scratch);
            _scratch = null;
        }
    }
}
