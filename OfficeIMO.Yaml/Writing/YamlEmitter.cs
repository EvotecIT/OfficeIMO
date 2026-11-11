namespace OfficeIMO.Yaml;

/// <summary>Where a fragment's single node goes when the model re-emits a node in place.</summary>
internal enum YamlSlot : byte {
    /// <summary>At a position where block collections continue on the current line: a sequence item, an explicit key or the document root.</summary>
    Inline,
    /// <summary>After a mapping key's colon: block collections start on the next line.</summary>
    MappingValue,
    /// <summary>Inside a flow collection.</summary>
    Flow
}

/// <summary>
/// The one emitter core behind <see cref="YamlWriter"/> and the model's save path. It writes events as YAML text with
/// a small frame stack, tracking the column so layout decisions (compact collections, explicit keys, block scalar
/// indentation) need no look-back. Block collections are opened lazily, so an empty one becomes <c>[]</c> or <c>{}</c>.
/// The model starts it in a fragment state (<see cref="BeginSlot"/>, <see cref="BeginEntries"/>) to emit a node or
/// entries into existing text. Writing allocates nothing per call; text goes straight to the <see cref="TextWriter"/>.
/// </summary>
internal sealed class YamlEmitter {
    private const string Spaces = "                                                                ";

    private enum FrameKind : byte { Document, BlockMapping, BlockSequence, FlowMapping, FlowSequence, Slot }

    private enum Shape : byte { Inline, Block, BlockScalar }

    // A frame holds no references, so pushing and copying one needs no GC write barriers; the properties of a pending
    // collection, rare, wait in the scratch's side arrays at the frame's index.
    private struct Frame {
        public FrameKind Kind;
        public YamlSlot Slot;
        /// <summary>Block mapping: key column. Block sequence: dash column. Slot: the enclosing collection's indentation.</summary>
        public int Indent;
        public int Count;
        public bool AtValue;
        public bool ExplicitKey;
        public bool PendingColon;
        public bool AliasKey;
        /// <summary>Spaces to write before the first entry, which continues the current line; -1 when entries start on new lines.</summary>
        public int Padding;
        /// <summary>A block collection whose layout is not decided until its first entry or its end.</summary>
        public bool Pending;
        /// <summary>A pending collection's anchor or tag waits in the scratch's side arrays.</summary>
        public bool HasProperties;
        /// <summary>Document: <c>---</c> was written. Slot: its node was written.</summary>
        public bool Marker;
        public bool HasRoot;
    }

    private readonly TextWriter _out;
    /// <summary>The output when it is a <see cref="YamlCharBuffer"/>, which is written without virtual calls.</summary>
    private readonly YamlCharBuffer? _buffer;
    private YamlWriterOptions _options = null!;
    private string _newLine = null!;
    // The options' numbers, read once per emitter instead of once per node.
    private int _mappingIndent;
    private int _dashOffset;
    private int _itemPadding;
    private const int LineStartCacheSize = 256;

    // A struct element, so filling the cache needs no covariance check.
    private struct LineStartEntry {
        public string? Text;
    }

    // Line break plus indentation per indentation, shared by every emitter with the same line break.
    private static readonly LineStartEntry[] s_lfLineStarts = new LineStartEntry[LineStartCacheSize];
    private static readonly LineStartEntry[] s_crlfLineStarts = new LineStartEntry[LineStartCacheSize];
    private LineStartEntry[] _lineStarts = null!;
    /// <summary>
    /// What an emitter reuses from the previous one on its thread: the frame stack, the side arrays for pending
    /// collection properties, the number formatting buffer and the key style table. One per thread is handed from
    /// emitter to emitter through <see cref="Release"/>, so a call allocates none of it and reaches thread-local storage once.
    /// </summary>
    private sealed class Scratch {
        internal Frame[] Frames = new Frame[8];
        internal string?[]? Anchors;
        internal string?[]? Tags;
        /// <summary>Room to format a number before it is copied out; nothing runs between formatting into it and copying from it.</summary>
        internal readonly char[] Number = new char[48];
        internal YamlKeyStyles.Entry[]? KeyStyles;
    }

    [ThreadStatic] private static Scratch? t_scratch;
    private const int MaxReusedFrames = 1024;
    private Scratch? _scratch;
    private Frame[] _frames = null!;
    // The deepest the stack has been, so releasing it clears only the frames used.
    private int _maxDepth;
    private int _depth;
    private int _column;
    private int _documents;
    private string? _anchor;
    private string? _tag;
    private string? _trailingComment;
    private string? _lastTag;
    private string? _lastTagText;
    private bool _lastTagRedefined;
#if !NET8_0_OR_GREATER
    private char[]? _copyScratch;
#endif

    public YamlEmitter(TextWriter output, YamlWriterOptions options) {
        _out = output;
        _buffer = output as YamlCharBuffer;
        Configure(options);
    }

    /// <summary>Prepares the emitter for a new stream with <paramref name="options"/>, taking this thread's reusable state.</summary>
    private void Configure(YamlWriterOptions options) {
        options.Validate();
        _options = options;
        _newLine = options.NewLine;
        _mappingIndent = options.MappingIndent;
        _dashOffset = options.SequenceDashOffset;
        _itemPadding = options.ItemPadding;
        // The options allow only "\n" and "\r\n".
        _lineStarts = _newLine.Length == 1 ? s_lfLineStarts : s_crlfLineStarts;
        Scratch scratch = t_scratch ?? new Scratch();
        t_scratch = null;
        _scratch = scratch;
        _frames = scratch.Frames;
    }

    /// <summary>Starts a new stream into the same output, after <see cref="Release"/>, as a new emitter would.</summary>
    public void Reset(YamlWriterOptions options) {
        _depth = 0;
        _maxDepth = 0;
        _column = 0;
        _documents = 0;
        _anchor = null;
        _tag = null;
        _trailingComment = null;
        _lastTag = null;
        _lastTagText = null;
        _lastTagRedefined = false;
        Configure(options);
    }

    /// <summary>Hands the reusable state to the next emitter on this thread. The emitter cannot be used afterwards.</summary>
    public void Release() {
        Scratch? scratch = _scratch;
        if (scratch is null) return;
        _scratch = null;
        Frame[] frames = _frames;
        _frames = Array.Empty<Frame>();
        _depth = 0;
        if (frames.Length > MaxReusedFrames) return;
        int used = Math.Min(_maxDepth, frames.Length);
        Array.Clear(frames, 0, used);
        if (scratch.Anchors is not null) {
            Array.Clear(scratch.Anchors, 0, Math.Min(used, scratch.Anchors.Length));
            Array.Clear(scratch.Tags!, 0, Math.Min(used, scratch.Tags!.Length));
        }
        scratch.Frames = frames;
        t_scratch = scratch;
    }

    /// <summary>A buffer to format a number into before writing it.</summary>
    public char[] NumberBuffer => GetScratch().Number;

    /// <summary>The default style of <paramref name="key"/> as a block mapping key; see <see cref="YamlKeyStyles"/>.</summary>
    public YamlStyle BlockKeyStyle(string key) => YamlKeyStyles.Block(key, ref GetScratch().KeyStyles);

    /// <summary>The emitter's reusable state; after <see cref="Release"/>, a new one.</summary>
    private Scratch GetScratch() {
        Scratch? scratch = _scratch;
        if (scratch is null) {
            scratch = new Scratch();
            _scratch = scratch;
        }
        return scratch;
    }

    public YamlWriterOptions Options => _options;

    /// <summary>The number of documents started so far.</summary>
    public int Documents => _documents;

    /// <summary>Whether the innermost open collection, or the slot being written, is a flow context.</summary>
    public bool InFlow {
        get {
            if (_depth == 0) return false;
            ref Frame top = ref _frames[_depth - 1];
            return top.Kind is FrameKind.FlowMapping or FrameKind.FlowSequence || (top.Kind == FrameKind.Slot && top.Slot == YamlSlot.Flow);
        }
    }

    /// <summary>Whether the next node is a mapping key.</summary>
    public bool AtKey {
        get {
            if (_depth == 0) return false;
            ref Frame top = ref _frames[_depth - 1];
            return (top.Kind is FrameKind.BlockMapping or FrameKind.FlowMapping) && !top.AtValue;
        }
    }

    /// <summary>Whether the stream is between documents, with no root written yet in the open one.</summary>
    public bool AtDocumentStart => _depth == 0 || (_depth == 1 && _frames[0].Kind == FrameKind.Document && !_frames[0].HasRoot);

    // ------------------------------------------------------------------ fragments

    /// <summary>Starts a fragment holding one node at <paramref name="column"/> of a line.</summary>
    public void BeginSlot(YamlSlot slot, int parentIndent, int column) {
        _depth = 0;
        Push(new Frame { Kind = FrameKind.Slot, Slot = slot, Indent = parentIndent, Padding = -1 });
        _column = column;
    }

    /// <summary>
    /// Starts a fragment of block entries at <paramref name="indent"/>. At column 0 the first entry starts its own line;
    /// at a column past the indentation it continues the current line (the first entry of a compact collection).
    /// </summary>
    public void BeginEntries(bool mapping, int indent, int column) {
        _depth = 0;
        Push(new Frame { Kind = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence, Indent = indent, Padding = column > 0 ? 0 : -1 });
        _column = column;
    }

    /// <summary>Ends a fragment, flushing a pending trailing comment. A line break is written only when <paramref name="endLine"/> is set.</summary>
    public void EndFragment(bool endLine) {
        if (endLine) {
            if (_column > 0) NewLine();
        } else if (_trailingComment is not null) {
            WriteTrailingComment();
        }
        _depth = 0;
    }

    /// <summary>Starts a new line at the indentation of the entries being written, after a fragment ends mid-line.</summary>
    public void ContinueLine(int indent) {
        NewLine();
        WriteSpaces(indent);
    }

    // ------------------------------------------------------------------ documents

    public void StartDocument(bool explicitMarker) {
        if (_depth > 1 || (_depth == 1 && _frames[0].Kind != FrameKind.Document)) throw Invalid("A document cannot start inside a collection.");
        if (_depth == 1 && !_frames[0].HasRoot && _frames[0].Marker) throw Invalid("The document has no root node yet.");
        if (_depth == 1) CloseDocument();
        EnsureLineStart();
        _depth = 0;
        var frame = new Frame { Kind = FrameKind.Document, Padding = -1 };
        if (explicitMarker || _documents > 0) {
            WriteRaw("---");
            frame.Marker = true;
        }
        _documents++;
        Push(frame);
    }

    public void EndDocument() {
        if (_depth != 1 || _frames[0].Kind != FrameKind.Document) throw Invalid("No document is open, or a collection is still open.");
        if (!_frames[0].HasRoot) throw Invalid("The document has no root node.");
        EnsureLineStart();
        WriteRaw("...");
        NewLine();
        _depth = 0;
    }

    /// <summary>Whether no collection is open, so the stream can be completed.</summary>
    public bool IsComplete => _depth <= 1;

    /// <summary>Completes the stream: closes the open document and ends the last line.</summary>
    public void Finish() {
        if (_depth > 1) throw Invalid("A collection is still open.");
        if (_trailingComment is not null || _column > 0) NewLine();
        _depth = 0;
    }

    private void CloseDocument() {
        if (_column > 0 || _trailingComment is not null) NewLine();
    }

    // ------------------------------------------------------------------ properties and comments

    public void SetAnchor(string anchor) {
        if (!YamlScalarFormat.IsValidAnchor(anchor)) throw new ArgumentException("'" + anchor + "' is not a valid anchor name: use printable characters without spaces or , [ ] { }.", nameof(anchor));
        if (_anchor is not null) throw Invalid("An anchor is already pending for the next node.");
        _anchor = anchor;
    }

    public void SetTag(string tag, bool secondaryHandleRedefined = false) {
        if (_tag is not null) throw Invalid("A tag is already pending for the next node.");
        // Repeated tags (a custom tag on every item) reuse their formatted text.
        if (!ReferenceEquals(tag, _lastTag) || secondaryHandleRedefined != _lastTagRedefined) {
            _lastTagText = YamlScalarFormat.FormatTag(tag, secondaryHandleRedefined);
            _lastTag = tag;
            _lastTagRedefined = secondaryHandleRedefined;
        }
        _tag = _lastTagText;
    }

    /// <summary>Writes comment lines at the indentation of the next entry. Each line of <paramref name="text"/> becomes one comment.</summary>
    public void WriteCommentLines(string text) {
        int indent = CommentIndent();
        int start = 0;
        while (true) {
            int end = text.IndexOf('\n', start);
            int stop = end < 0 ? text.Length : end;
            int length = stop > start && text[stop - 1] == '\r' ? stop - start - 1 : stop - start;
            EnsureLineStart();
            WriteSpaces(indent);
            Write('#');
            if (length > 0) {
                Write(' ');
                WriteRun(text, start, length);
            }
            if (end < 0) break;
            start = end + 1;
        }
    }

    /// <summary>Writes an empty line before the next entry.</summary>
    public void WriteBlankLine() {
        CommentIndent();
        EnsureLineStart();
        NewLine();
    }

    /// <summary>Holds a comment for the end of the current line; it is written before the next line break.</summary>
    public void SetTrailingComment(string? text) => _trailingComment = text;

    private int CommentIndent() {
        if (_depth == 0) return 0;
        ref Frame top = ref _frames[_depth - 1];
        if (top.Pending) ResolvePending(empty: false);
        top = ref _frames[_depth - 1];
        switch (top.Kind) {
            case FrameKind.Document:
                return 0;
            case FrameKind.BlockMapping:
                if (top.AtValue) throw Invalid("A comment line cannot be written between a key and its value.");
                top.Padding = -1;
                return top.Indent;
            case FrameKind.BlockSequence:
                top.Padding = -1;
                return top.Indent;
            default:
                throw Invalid("A comment line cannot be written inside a flow collection or a single-node fragment.");
        }
    }

    // ------------------------------------------------------------------ nodes

    /// <summary>Writes a scalar in a style decided by the caller (see <see cref="ChooseStyle"/>).</summary>
    public void WriteScalar(string value, YamlStyle style) {
        bool empty = value.Length == 0 && style == YamlStyle.Plain;
        if (style is YamlStyle.Literal or YamlStyle.Folded) {
            int parentIndent = BeginNode(Shape.BlockScalar, empty: false, simpleKey: false, aliasKey: false, sequence: false, out int contentIndent, out _);
            WriteBlockScalar(value, style == YamlStyle.Folded, parentIndent, contentIndent);
            return;
        }
        // An empty key is written as a bare ':' (`: value`), which every reader takes as an empty key.
        bool simpleKey = value.IndexOf('\n') < 0;
        if (value.Length > YamlScalarFormat.MaxImplicitKeyLength) simpleKey = false;
        BeginNode(Shape.Inline, empty, simpleKey, aliasKey: false, sequence: false, out _, out _);
        switch (style) {
            case YamlStyle.Plain:
                if (value.IndexOf('\n') < 0) WriteRun(value, 0, value.Length);
                else WritePlainLines(value);
                break;
            case YamlStyle.SingleQuoted:
                WriteSingleQuoted(value);
                break;
            default:
                WriteDoubleQuoted(value);
                break;
        }
        EndInline();
    }

    /// <summary>
    /// Writes a plain or single-quoted scalar the caller knows has no line break and fits an implicit key, as
    /// <see cref="WriteScalar"/> would, without scanning it for breaks.
    /// </summary>
    public void WriteSingleLine(string value, YamlStyle style) {
        if (style == YamlStyle.Plain && value.Length > 0 && TryWriteBlockKey(value)) return;
        BeginNode(Shape.Inline, empty: value.Length == 0 && style == YamlStyle.Plain, simpleKey: true, aliasKey: false, sequence: false, out _, out _);
        if (style == YamlStyle.Plain) WriteRun(value, 0, value.Length);
        else WriteSingleQuoted(value);
        EndInline();
    }

    /// <summary>
    /// Writes a non-empty plain key of a block mapping, with no anchor or tag, its line start and colon into the char
    /// buffer in one step, leaving the text and the mapping's state that <see cref="BeginNode"/>, the key and
    /// <see cref="EndInline"/> leave. Returns false, having written nothing, anywhere else.
    /// </summary>
    private bool TryWriteBlockKey(string key) {
        if (_depth == 0 || _buffer is null || _anchor is not null || _tag is not null || _trailingComment is not null) return false;
        ref Frame top = ref _frames[_depth - 1];
        if (top.Kind != FrameKind.BlockMapping || top.AtValue) return false;
        if (top.Pending) {
            // The first key of a mapping just started lays the mapping out first, as BeginNode would.
            ResolvePending(empty: false);
            top = ref _frames[_depth - 1];
        }
        if (top.Padding < 0) {
            if (_column == 0 || (uint)top.Indent >= LineStartCacheSize) return false;
            _buffer.AppendKey(LineStart(top.Indent).AsSpan(), key);
            _column = top.Indent + key.Length + 1;
        } else {
            if (top.Padding > Spaces.Length) return false;
            _buffer.AppendKey(Spaces.AsSpan(0, top.Padding), key);
            _column += top.Padding + key.Length + 1;
            top.Padding = -1;
        }
        top.AtValue = true;
        top.ExplicitKey = false;
        top.PendingColon = false;
        top.AliasKey = false;
        return true;
    }

    /// <summary>Writes a plain scalar from a span without allocating, for typed values already in canonical form.</summary>
    public void WritePlain(ReadOnlySpan<char> value) {
        BeginNode(Shape.Inline, empty: value.Length == 0, simpleKey: value.Length > 0, aliasKey: false, sequence: false, out _, out _);
        WriteSpan(value);
        EndInline();
    }

    public void WriteAlias(string name) {
        if (!YamlScalarFormat.IsValidAnchor(name)) throw new ArgumentException("'" + name + "' is not a valid anchor name.", nameof(name));
        if (_anchor is not null || _tag is not null) throw Invalid("An alias cannot have an anchor or a tag.");
        BeginNode(Shape.Inline, empty: false, simpleKey: true, aliasKey: true, sequence: false, out _, out _);
        Write('*');
        WriteRun(name, 0, name.Length);
        EndInline();
    }

    /// <summary>
    /// Starts a block collection the caller knows will get at least one entry. Where its first entry would lay it out
    /// through the usual nesting (the value of a simple key or a sequence item, without properties) the layout is
    /// decided now, with the same text and state, so no pending collection waits for the entry.
    /// </summary>
    public void StartNonEmptyCollection(bool mapping) {
        if (_depth > 0 && _anchor is null && _tag is null && _trailingComment is null) {
            ref Frame parent = ref _frames[_depth - 1];
            if (!parent.Pending) {
                if (parent.Kind == FrameKind.BlockMapping && parent.AtValue && !parent.ExplicitKey) {
                    parent.AtValue = false;
                    parent.Count++;
                    Push(new Frame { Kind = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence, Indent = parent.Indent + (mapping ? _mappingIndent : _dashOffset), Padding = -1 });
                    return;
                }
                if (parent.Kind == FrameKind.BlockSequence) {
                    StartEntry(ref parent);
                    Write('-');
                    parent.Count++;
                    Push(new Frame { Kind = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence, Indent = _column + _itemPadding, Padding = _itemPadding });
                    return;
                }
            }
        }
        StartCollection(mapping, YamlStyle.Block);
    }

    public void StartCollection(bool mapping, YamlStyle style) {
        if (style == YamlStyle.Flow || InFlow) {
            BeginNode(Shape.Inline, empty: false, simpleKey: false, aliasKey: false, sequence: !mapping, out _, out _);
            Write(mapping ? '{' : '[');
            Push(new Frame { Kind = mapping ? FrameKind.FlowMapping : FrameKind.FlowSequence, Padding = -1 });
            return;
        }
        if (_depth == 0) StartDocument(explicitMarker: false);
        else if (_frames[_depth - 1].Pending) ResolvePending(empty: false);
        bool properties = _anchor is not null || _tag is not null;
        if (properties) {
            SetFrameProperties(_depth, _anchor, _tag);
            _anchor = null;
            _tag = null;
        }
        Push(new Frame { Kind = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence, Pending = true, HasProperties = properties, Padding = -1 });
    }

    private void SetFrameProperties(int index, string? anchor, string? tag) {
        Scratch scratch = GetScratch();
        if (scratch.Anchors is null || index >= scratch.Anchors.Length) {
            int size = Math.Max(8, (index + 1) * 2);
            Array.Resize(ref scratch.Anchors, size);
            Array.Resize(ref scratch.Tags, size);
        }
        scratch.Anchors[index] = anchor;
        scratch.Tags![index] = tag;
    }

    public void EndCollection(bool mapping) {
        if (_depth == 0) throw Invalid("No collection is open.");
        ref Frame top = ref _frames[_depth - 1];
        FrameKind expectedBlock = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence;
        FrameKind expectedFlow = mapping ? FrameKind.FlowMapping : FrameKind.FlowSequence;
        if (top.Kind != expectedBlock && top.Kind != expectedFlow) throw Invalid("The innermost open collection is not a " + (mapping ? "mapping." : "sequence."));
        if (top.AtValue) throw Invalid("The mapping has a key without a value.");
        if (top.Kind == expectedFlow) {
            _depth--;
            Write(mapping ? '}' : ']');
            EndInline();
            return;
        }
        if (top.Pending) {
            ResolvePending(empty: true);
            return;
        }
        if (top.Count == 0) {
            // A comment opened the collection as a block; an empty flow collection still has to follow.
            int indent = top.Indent;
            _depth--;
            EnsureLineStart();
            WriteSpaces(indent);
            WriteRaw(mapping ? "{}" : "[]");
            return;
        }
        _depth--;
    }

    /// <summary>Picks a string's style for the next node's position; see <see cref="YamlScalarFormat.Choose"/>.</summary>
    public YamlStyle ChooseStyle(string value, YamlStyle requested, YamlStyle quote = YamlStyle.SingleQuoted) =>
        YamlScalarFormat.Choose(value, requested, InFlow, AtKey, quote);

    /// <summary>
    /// Whether the next node is a key of a block mapping, including the first key of one just started, with no
    /// anchor or tag set for it, where <see cref="WriteEntry"/> can write a whole entry.
    /// </summary>
    public bool AtBlockKey {
        get {
            if (_depth == 0 || _anchor is not null || _tag is not null) return false;
            ref Frame top = ref _frames[_depth - 1];
            return top.Kind == FrameKind.BlockMapping && !top.AtValue;
        }
    }

    /// <summary>
    /// Writes a block mapping entry at <see cref="AtBlockKey"/>: a key and a value the caller knows are plain on one line.
    /// The text and the mapping's state are what a key and a value written one by one leave.
    /// </summary>
    public void WriteEntry(string key, ReadOnlySpan<char> value) {
        ref Frame top = ref _frames[_depth - 1];
        if (top.Pending) {
            ResolvePending(empty: false);
            top = ref _frames[_depth - 1];
        }
        if (_buffer is not null && value.Length > 0 && _trailingComment is null) {
            // The usual entry, on a line of its own or continuing a sequence item's line, goes into the char buffer in
            // one step, as StartLine or the padding and the pieces below would write it.
            if (top.Padding < 0) {
                if (_column > 0 && (uint)top.Indent < LineStartCacheSize) {
                    _buffer.AppendEntry(LineStart(top.Indent).AsSpan(), key, value);
                    _column = top.Indent + key.Length + 2 + value.Length;
                    top.ExplicitKey = false;
                    top.AliasKey = false;
                    top.Count++;
                    return;
                }
            } else if (top.Padding <= Spaces.Length) {
                _buffer.AppendEntry(Spaces.AsSpan(0, top.Padding), key, value);
                _column += top.Padding + key.Length + 2 + value.Length;
                top.Padding = -1;
                top.ExplicitKey = false;
                top.AliasKey = false;
                top.Count++;
                return;
            }
        }
        if (top.Padding >= 0) {
            WriteSpaces(top.Padding);
            top.Padding = -1;
        } else {
            StartLine(top.Indent);
        }
        WriteRun(key, 0, key.Length);
        if (_buffer is not null && value.Length > 0) {
            _buffer.Append(": ".AsSpan());
            _buffer.Append(value);
            _column += 2 + value.Length;
        } else {
            Write(':');
            if (value.Length > 0) {
                Write(' ');
                WriteSpan(value);
            }
        }
        top.ExplicitKey = false;
        top.AliasKey = false;
        top.Count++;
    }

    /// <summary>Whether the next node is an item of a block sequence, including the first of one just started, with no anchor or tag set for it.</summary>
    public bool AtBlockItem {
        get {
            if (_depth == 0 || _anchor is not null || _tag is not null) return false;
            return _frames[_depth - 1].Kind == FrameKind.BlockSequence;
        }
    }

    /// <summary>
    /// Writes a block sequence item at <see cref="AtBlockItem"/> whose text the caller knows is plain on one line, with
    /// the same text and sequence state as <see cref="WriteSingleLine"/> leaves.
    /// </summary>
    public void WriteItem(ReadOnlySpan<char> value) {
        ref Frame top = ref _frames[_depth - 1];
        if (top.Pending) {
            ResolvePending(empty: false);
            top = ref _frames[_depth - 1];
        }
        int padding = _itemPadding;
        if (_buffer is not null && value.Length > 0 && _trailingComment is null && padding <= Spaces.Length) {
            // The usual item, on a line of its own or continuing the line of an enclosing item, goes into the char
            // buffer in one step, as StartEntry and the pieces below would write it.
            if (top.Padding < 0) {
                if (_column > 0 && (uint)top.Indent < LineStartCacheSize) {
                    _buffer.AppendItem(LineStart(top.Indent).AsSpan(), Spaces.AsSpan(0, padding), value);
                    _column = top.Indent + 1 + padding + value.Length;
                    top.Count++;
                    return;
                }
            } else if (top.Padding <= Spaces.Length) {
                _buffer.AppendItem(Spaces.AsSpan(0, top.Padding), Spaces.AsSpan(0, padding), value);
                _column += top.Padding + 1 + padding + value.Length;
                top.Padding = -1;
                top.Count++;
                return;
            }
        }
        StartEntry(ref top);
        Write('-');
        top.Count++;
        if (value.Length > 0) {
            WriteSpaces(padding);
            WriteSpan(value);
        }
    }

    // ------------------------------------------------------------------ layout

    /// <summary>
    /// Writes what precedes a node in its parent (line break, indentation, dash, key colon, separators) and the node's
    /// anchor and tag. For block collections it returns the child indentation and the padding before a first entry that
    /// continues the current line (-1 when entries start on new lines); for block scalars it returns the parent's
    /// indentation and the content indentation.
    /// </summary>
    private int BeginNode(Shape shape, bool empty, bool simpleKey, bool aliasKey, bool sequence, out int childIndent, out int padding) {
        childIndent = 0;
        padding = -1;
        if (_depth == 0) StartDocument(explicitMarker: false);
        ref Frame top = ref _frames[_depth - 1];
        // The two most common nodes, an inline key or value of a block mapping without properties, skip the general
        // layout below; the text they write is the same.
        if (top.Kind == FrameKind.BlockMapping && shape == Shape.Inline && !top.Pending && _anchor is null && _tag is null) {
            if (top.AtValue) {
                if (!top.ExplicitKey) {
                    top.AtValue = false;
                    top.Count++;
                    childIndent = top.Indent + _mappingIndent;
                    if (!empty) Write(' ');
                    return top.Indent;
                }
            } else if (simpleKey && top.Padding < 0) {
                StartLine(top.Indent);
                top.AtValue = true;
                top.ExplicitKey = false;
                top.PendingColon = true;
                top.AliasKey = aliasKey;
                return top.Indent;
            }
        }
        if (top.Pending) {
            ResolvePending(empty: false);
            top = ref _frames[_depth - 1];
        }
        string? anchor = _anchor, tag = _tag;
        _anchor = null;
        _tag = null;
        bool props = anchor is not null || tag is not null;
        int itemPadding = _itemPadding;
        switch (top.Kind) {
            case FrameKind.Document:
                if (top.HasRoot) throw Invalid("The document already has a root node; start a new document first.");
                top.HasRoot = true;
                if (top.Marker && (props || shape != Shape.Block)) Write(' ');
                if (shape == Shape.Block) {
                    WriteProperties(anchor, tag, trailingSpace: false);
                    childIndent = 0;
                    return 0;
                }
                // An empty root without properties or a marker would leave no document at all.
                if (empty && !props && !top.Marker) {
                    WriteRaw("---");
                    top.Marker = true;
                }
                WriteProperties(anchor, tag, trailingSpace: !empty || shape == Shape.BlockScalar);
                if (shape == Shape.BlockScalar) childIndent = _mappingIndent;
                return 0;
            case FrameKind.BlockSequence: {
                StartEntry(ref top);
                Write('-');
                top.Count++;
                return AfterIndicator(shape, empty, anchor, tag, props, top.Indent, top.Indent + 1 + itemPadding, itemPadding, out childIndent, out padding);
            }
            case FrameKind.BlockMapping:
                if (!top.AtValue) {
                    StartEntry(ref top);
                    top.AtValue = true;
                    if (shape != Shape.Inline || !simpleKey) {
                        top.ExplicitKey = true;
                        Write('?');
                        return AfterIndicator(shape, empty, anchor, tag, props, top.Indent, top.Indent + 2, 1, out childIndent, out padding);
                    }
                    top.ExplicitKey = false;
                    top.PendingColon = true;
                    top.AliasKey = aliasKey;
                    WriteProperties(anchor, tag, trailingSpace: true);
                    return top.Indent;
                }
                top.AtValue = false;
                top.Count++;
                if (top.ExplicitKey) {
                    EnsureLineStart();
                    WriteSpaces(top.Indent);
                    Write(':');
                    return AfterIndicator(shape, empty, anchor, tag, props, top.Indent, top.Indent + 2, 1, out childIndent, out padding);
                }
                return AfterColon(shape, empty, anchor, tag, props, top.Indent, sequence, out childIndent);
            case FrameKind.FlowSequence:
                if (shape != Shape.Inline) throw Invalid("Block nodes cannot be written inside a flow collection.");
                if (top.Count > 0) WriteRaw(", ");
                top.Count++;
                WriteProperties(anchor, tag, trailingSpace: true);
                return 0;
            case FrameKind.FlowMapping:
                if (shape != Shape.Inline) throw Invalid("Block nodes cannot be written inside a flow collection.");
                if (!top.AtValue) {
                    if (top.Count > 0) WriteRaw(", ");
                    top.AtValue = true;
                    top.AliasKey = aliasKey;
                    top.ExplicitKey = empty && !props;
                    if (!simpleKey && !empty) WriteRaw("? ");
                    WriteProperties(anchor, tag, trailingSpace: true);
                    return 0;
                }
                top.AtValue = false;
                top.Count++;
                // An entry with an empty key keeps its ':', or nothing of it would remain.
                if (!empty || props || top.ExplicitKey) WriteRaw(top.AliasKey ? " : " : ": ");
                WriteProperties(anchor, tag, trailingSpace: true);
                return 0;
            default: // Slot
                if (top.Marker) throw Invalid("The fragment already holds its node.");
                top.Marker = true;
                switch (top.Slot) {
                    case YamlSlot.MappingValue:
                        return AfterColon(shape, empty, anchor, tag, props, top.Indent, sequence, out childIndent);
                    case YamlSlot.Flow:
                        if (shape != Shape.Inline) throw Invalid("Block nodes cannot be written inside a flow collection.");
                        WriteProperties(anchor, tag, trailingSpace: !empty);
                        return 0;
                    default:
                        if (shape == Shape.Block) {
                            // Without properties the first entry continues here; with them, entries align under the node.
                            childIndent = _column;
                            padding = props ? -1 : 0;
                            WriteProperties(anchor, tag, trailingSpace: false);
                            return top.Indent;
                        }
                        int nodeColumn = _column;
                        WriteProperties(anchor, tag, trailingSpace: !empty || shape == Shape.BlockScalar);
                        childIndent = nodeColumn > top.Indent ? nodeColumn : top.Indent + _mappingIndent;
                        return top.Indent;
                }
        }
    }

    /// <summary>Layout after <c>-</c>, <c>?</c> or an explicit <c>:</c>: inline nodes follow a space, block collections are compact.</summary>
    private int AfterIndicator(Shape shape, bool empty, string? anchor, string? tag, bool props, int parentIndent, int contentIndent, int indicatorPadding, out int childIndent, out int padding) {
        padding = -1;
        childIndent = contentIndent;
        switch (shape) {
            case Shape.Block:
                if (props) {
                    Write(' ');
                    WriteProperties(anchor, tag, trailingSpace: false);
                } else {
                    padding = indicatorPadding;
                }
                return parentIndent;
            case Shape.BlockScalar:
                WriteSpaces(indicatorPadding);
                WriteProperties(anchor, tag, trailingSpace: true);
                return parentIndent;
            default:
                if (props || !empty) {
                    WriteSpaces(indicatorPadding);
                    WriteProperties(anchor, tag, trailingSpace: !empty);
                }
                return parentIndent;
        }
    }

    /// <summary>Layout after a simple key's colon: inline nodes follow a space, block collections start on the next line.</summary>
    private int AfterColon(Shape shape, bool empty, string? anchor, string? tag, bool props, int keyIndent, bool sequence, out int childIndent) {
        childIndent = keyIndent + _mappingIndent;
        switch (shape) {
            case Shape.Block:
                if (props) {
                    Write(' ');
                    WriteProperties(anchor, tag, trailingSpace: false);
                }
                // A sequence with properties on the key line is indented, so its dash is not read as part of the key.
                if (sequence) childIndent = keyIndent + (props && _dashOffset == 0 ? _mappingIndent : _dashOffset);
                return keyIndent;
            case Shape.BlockScalar:
                Write(' ');
                WriteProperties(anchor, tag, trailingSpace: true);
                return keyIndent;
            default:
                if (props || !empty) {
                    Write(' ');
                    WriteProperties(anchor, tag, trailingSpace: !empty);
                }
                return keyIndent;
        }
    }

    private void StartEntry(ref Frame top) {
        if (top.Padding >= 0) {
            WriteSpaces(top.Padding);
            top.Padding = -1;
        } else {
            StartLine(top.Indent);
        }
    }

    /// <summary>Called after an inline node: a simple block key gets its colon.</summary>
    private void EndInline() {
        if (_depth == 0) return;
        ref Frame top = ref _frames[_depth - 1];
        if (top.Kind == FrameKind.BlockMapping && top.PendingColon) {
            top.PendingColon = false;
            WriteRaw(top.AliasKey ? " :" : ":");
        }
    }

    /// <summary>Decides the layout of a pending block collection, at its first entry or, when empty, at its end.</summary>
    private void ResolvePending(bool empty) {
        if (!empty && _depth >= 2 && !_frames[_depth - 1].HasProperties && _trailingComment is null) {
            // The usual nesting, a collection without properties as the value of a simple key or as a sequence item,
            // lays out as BeginNode would below, without the general layout.
            ref Frame pending = ref _frames[_depth - 1];
            ref Frame parent = ref _frames[_depth - 2];
            if (parent.Kind == FrameKind.BlockMapping && parent.AtValue && !parent.ExplicitKey) {
                parent.AtValue = false;
                parent.Count++;
                pending.Indent = parent.Indent + (pending.Kind == FrameKind.BlockSequence ? _dashOffset : _mappingIndent);
                pending.Padding = -1;
                pending.Pending = false;
                return;
            }
            if (parent.Kind == FrameKind.BlockSequence) {
                StartEntry(ref parent);
                Write('-');
                parent.Count++;
                int itemPadding = _itemPadding;
                pending.Indent = _column + itemPadding;
                pending.Padding = itemPadding;
                pending.Pending = false;
                return;
            }
        }
        Frame frame = _frames[_depth - 1];
        _depth--;
        // The first entry's own anchor and tag may already be pending; the collection's go first.
        string? entryAnchor = _anchor, entryTag = _tag;
        _anchor = null;
        _tag = null;
        if (frame.HasProperties) {
            Scratch scratch = _scratch!;
            _anchor = scratch.Anchors![_depth];
            _tag = scratch.Tags![_depth];
            scratch.Anchors[_depth] = null;
            scratch.Tags[_depth] = null;
        }
        bool mapping = frame.Kind == FrameKind.BlockMapping;
        if (empty) {
            BeginNode(Shape.Inline, empty: false, simpleKey: false, aliasKey: false, sequence: !mapping, out _, out _);
            WriteRaw(mapping ? "{}" : "[]");
            EndInline();
            _anchor = entryAnchor;
            _tag = entryTag;
            return;
        }
        BeginNode(Shape.Block, empty: false, simpleKey: false, aliasKey: false, sequence: !mapping, out int childIndent, out int padding);
        _anchor = entryAnchor;
        _tag = entryTag;
        frame.Pending = false;
        frame.HasProperties = false;
        frame.Indent = childIndent;
        frame.Padding = padding;
        if (padding >= 0) frame.Indent = _column + padding;
        _frames[_depth++] = frame;
    }

    // ------------------------------------------------------------------ scalar text

    private void WriteBlockScalar(string value, bool folded, int parentIndent, int contentIndent) {
        int trailing = 0;
        while (trailing < value.Length && value[value.Length - 1 - trailing] == '\n') trailing++;
        int bodyLength = value.Length - trailing;
        Write(folded ? '>' : '|');
        // An explicit indentation indicator is needed when the content starts with a space or a line break.
        if (value.Length > 0 && (value[0] == ' ' || value[0] == '\n')) {
            int indicator = contentIndent - Math.Max(parentIndent, 0);
            Write((char)('0' + indicator));
        }
        if (trailing == 0) Write('-');
        else if (trailing > 1 || bodyLength == 0) Write('+');
        if (_trailingComment is not null) WriteTrailingComment();
        if (bodyLength == 0) {
            NewLine();
            for (int i = 0; i < trailing; i++) NewLine();
            return;
        }
        int start = 0;
        int emptyRun = 0;
        bool previousNormal = false, anyContent = false;
        while (true) {
            int end = value.IndexOf('\n', start, bodyLength - start);
            if (end < 0) end = bodyLength;
            if (end == start) {
                emptyRun++;
            } else {
                bool normal = value[start] != ' ' && value[start] != '\t';
                // Folding turns the breaks between two unindented lines into one fewer: write one more.
                int breaks = emptyRun + 1 + (folded && anyContent && previousNormal && normal ? 1 : 0);
                for (int i = 0; i < breaks; i++) NewLine();
                WriteSpaces(contentIndent);
                WriteRun(value, start, end - start);
                emptyRun = 0;
                previousNormal = normal;
                anyContent = true;
            }
            if (end == bodyLength) break;
            start = end + 1;
        }
        NewLine();
        for (int i = 1; i < trailing; i++) NewLine();
    }

    /// <summary>
    /// Writes a plain scalar over several lines. Each line break of the value becomes an empty line, because folding
    /// turns a single break into a space; continuation lines are indented past the enclosing block level.
    /// </summary>
    private void WritePlainLines(string value) {
        int indent = BlockLevel() + _mappingIndent;
        int start = 0;
        bool first = true;
        while (true) {
            int end = value.IndexOf('\n', start);
            int stop = end < 0 ? value.Length : end;
            if (stop > start) {
                if (!first) WriteSpaces(indent);
                WriteRun(value, start, stop - start);
                first = false;
            }
            if (end < 0) break;
            NewLine();
            if (stop > start) NewLine();
            start = end + 1;
        }
    }

    /// <summary>The indentation of the innermost block level, which continuation lines must exceed.</summary>
    private int BlockLevel() {
        for (int i = _depth - 1; i >= 0; i--) {
            switch (_frames[i].Kind) {
                case FrameKind.BlockMapping:
                case FrameKind.BlockSequence:
                case FrameKind.Slot:
                    return Math.Max(_frames[i].Indent, 0);
                case FrameKind.Document:
                    return 0;
            }
        }
        return 0;
    }

    private void WriteSingleQuoted(string value) {
        Write('\'');
        int start = 0;
        int quote;
        // Each quote is doubled: write through it, then one more.
        while ((quote = value.IndexOf('\'', start)) >= 0) {
            WriteRun(value, start, quote - start + 1);
            Write('\'');
            start = quote + 1;
        }
        WriteRun(value, start, value.Length - start);
        Write('\'');
    }

    private void WriteDoubleQuoted(string value) {
        Write('"');
        int start = 0;
        for (int i = 0; i < value.Length; i++) {
            char c = value[i];
            string? escape = c switch {
                '"' => "\\\"",
                '\\' => "\\\\",
                '\0' => "\\0",
                '\a' => "\\a",
                '\b' => "\\b",
                '\t' => "\\t",
                '\n' => "\\n",
                '\v' => "\\v",
                '\f' => "\\f",
                '\r' => "\\r",
                '\u001B' => "\\e",
                '\u0085' => "\\N",
                '\u2028' => "\\L",
                '\u2029' => "\\P",
                _ => null
            };
            bool hex = escape is null && NeedsHexEscape(value, i);
            if (escape is null && !hex) continue;
            WriteRun(value, start, i - start);
            if (escape is not null) {
                WriteRaw(escape);
            } else if (c < 0x100) {
                WriteRaw("\\x");
                WriteHex(c, 2);
            } else {
                WriteRaw("\\u");
                WriteHex(c, 4);
            }
            start = i + 1;
        }
        WriteRun(value, start, value.Length - start);
        Write('"');
    }

    private static bool NeedsHexEscape(string value, int i) {
        char c = value[i];
        if (char.IsHighSurrogate(c)) return i + 1 >= value.Length || !char.IsLowSurrogate(value[i + 1]);
        if (char.IsLowSurrogate(c)) return i == 0 || !char.IsHighSurrogate(value[i - 1]);
        return !YamlScalarFormat.IsPrintableInline(c);
    }

    private void WriteHex(int value, int digits) {
        for (int shift = (digits - 1) * 4; shift >= 0; shift -= 4) Write("0123456789ABCDEF"[(value >> shift) & 15]);
    }

    private void WriteProperties(string? anchor, string? tag, bool trailingSpace) {
        if (anchor is null && tag is null) return;
        if (anchor is not null) {
            Write('&');
            WriteRaw(anchor);
        }
        if (tag is not null) {
            if (anchor is not null) Write(' ');
            WriteRaw(tag);
        }
        if (trailingSpace) Write(' ');
    }

    // ------------------------------------------------------------------ output primitives

    private void EnsureLineStart() {
        if (_column > 0 || _trailingComment is not null) NewLine();
    }

    /// <summary>
    /// Ends the current line, if any, and indents the next one. Into the char buffer, the line break and the
    /// indentation go as one run cached per indentation.
    /// </summary>
    private void StartLine(int indent) {
        if (_column > 0 && _trailingComment is null && _buffer is not null && (uint)indent < LineStartCacheSize) {
            _buffer.Append(LineStart(indent).AsSpan());
            _column = indent;
            return;
        }
        EnsureLineStart();
        WriteSpaces(indent);
    }

    /// <summary>The line break and indentation as one string, cached per indentation below <see cref="LineStartCacheSize"/>.</summary>
    private string LineStart(int indent) {
        ref LineStartEntry entry = ref _lineStarts[indent];
        return entry.Text ??= _newLine + new string(' ', indent);
    }

    private void NewLine() {
        if (_trailingComment is not null) WriteTrailingComment();
        if (_buffer is not null) _buffer.Append(_newLine.AsSpan());
        else _out.Write(_newLine);
        _column = 0;
    }

    private void WriteTrailingComment() {
        string text = _trailingComment!;
        _trailingComment = null;
        WriteRaw(_column > 0 ? " #" : "#");
        if (text.Length > 0) {
            Write(' ');
            WriteRaw(text);
        }
    }

    private void WriteSpaces(int count) {
        if (_buffer is not null && count > Spaces.Length) {
            _buffer.AppendSpaces(count);
            _column += count;
            return;
        }
        while (count > 0) {
            int n = Math.Min(count, Spaces.Length);
            WriteRun(Spaces, 0, n);
            count -= n;
        }
    }

    private void Write(char c) {
        if (_buffer is not null) _buffer.Write(c);
        else _out.Write(c);
        _column++;
    }

    private void WriteRaw(string text) => WriteRun(text, 0, text.Length);

    private void WriteRun(string text, int start, int length) {
        if (length <= 0) return;
        if (_buffer is not null) {
            _buffer.Append(text.AsSpan(start, length));
            _column += length;
            return;
        }
#if NET8_0_OR_GREATER
        _out.Write(text.AsSpan(start, length));
#else
        char[] scratch = _copyScratch ??= new char[256];
        int offset = 0;
        while (offset < length) {
            int n = Math.Min(scratch.Length, length - offset);
            text.CopyTo(start + offset, scratch, 0, n);
            _out.Write(scratch, 0, n);
            offset += n;
        }
#endif
        _column += length;
    }

    private void WriteSpan(ReadOnlySpan<char> text) {
        if (text.Length == 0) return;
        if (_buffer is not null) {
            _buffer.Append(text);
            _column += text.Length;
            return;
        }
#if NET8_0_OR_GREATER
        _out.Write(text);
#else
        char[] scratch = _copyScratch ??= new char[256];
        int offset = 0;
        while (offset < text.Length) {
            int n = Math.Min(scratch.Length, text.Length - offset);
            text.Slice(offset, n).CopyTo(scratch);
            _out.Write(scratch, 0, n);
            offset += n;
        }
#endif
        _column += text.Length;
    }

    private void Push(Frame frame) {
        if (_depth == _frames.Length) Array.Resize(ref _frames, Math.Max(8, _frames.Length * 2));
        _frames[_depth++] = frame;
        if (_depth > _maxDepth) _maxDepth = _depth;
    }

    private static InvalidOperationException Invalid(string message) => new InvalidOperationException(message);
}
