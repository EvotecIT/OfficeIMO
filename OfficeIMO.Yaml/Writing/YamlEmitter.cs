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
        public string? Anchor;
        public string? Tag;
        /// <summary>Document: <c>---</c> was written. Slot: its node was written.</summary>
        public bool Marker;
        public bool HasRoot;
    }

    private readonly TextWriter _out;
    /// <summary>The output when it is a <see cref="YamlCharBuffer"/>, which is written without virtual calls.</summary>
    private readonly YamlCharBuffer? _buffer;
    private readonly YamlWriterOptions _options;
    private readonly string _newLine;
    private Frame[] _frames = new Frame[16];
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
    private char[]? _scratch;
#endif

    public YamlEmitter(TextWriter output, YamlWriterOptions options) {
        _out = output;
        _buffer = output as YamlCharBuffer;
        _options = options;
        _newLine = options.NewLine;
        options.Validate();
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
        BeginNode(Shape.Inline, empty: value.Length == 0 && style == YamlStyle.Plain, simpleKey: true, aliasKey: false, sequence: false, out _, out _);
        if (style == YamlStyle.Plain) WriteRun(value, 0, value.Length);
        else WriteSingleQuoted(value);
        EndInline();
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

    public void StartCollection(bool mapping, YamlStyle style) {
        if (style == YamlStyle.Flow || InFlow) {
            BeginNode(Shape.Inline, empty: false, simpleKey: false, aliasKey: false, sequence: !mapping, out _, out _);
            Write(mapping ? '{' : '[');
            Push(new Frame { Kind = mapping ? FrameKind.FlowMapping : FrameKind.FlowSequence, Padding = -1 });
            return;
        }
        if (_depth == 0) StartDocument(explicitMarker: false);
        else if (_frames[_depth - 1].Pending) ResolvePending(empty: false);
        Push(new Frame { Kind = mapping ? FrameKind.BlockMapping : FrameKind.BlockSequence, Pending = true, Anchor = _anchor, Tag = _tag, Padding = -1 });
        _anchor = null;
        _tag = null;
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
                    childIndent = top.Indent + _options.MappingIndent;
                    if (!empty) Write(' ');
                    return top.Indent;
                }
            } else if (simpleKey && top.Padding < 0) {
                EnsureLineStart();
                WriteSpaces(top.Indent);
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
        int itemPadding = _options.ItemPadding;
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
                if (shape == Shape.BlockScalar) childIndent = _options.MappingIndent;
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
                        childIndent = nodeColumn > top.Indent ? nodeColumn : top.Indent + _options.MappingIndent;
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
        childIndent = keyIndent + _options.MappingIndent;
        switch (shape) {
            case Shape.Block:
                if (props) {
                    Write(' ');
                    WriteProperties(anchor, tag, trailingSpace: false);
                }
                // A sequence with properties on the key line is indented, so its dash is not read as part of the key.
                if (sequence) childIndent = keyIndent + (props && _options.SequenceDashOffset == 0 ? _options.MappingIndent : _options.SequenceDashOffset);
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
            EnsureLineStart();
            WriteSpaces(top.Indent);
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
        Frame frame = _frames[_depth - 1];
        _depth--;
        // The first entry's own anchor and tag may already be pending; the collection's go first.
        string? entryAnchor = _anchor, entryTag = _tag;
        _anchor = frame.Anchor;
        _tag = frame.Tag;
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
        frame.Anchor = null;
        frame.Tag = null;
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
        int indent = BlockLevel() + _options.MappingIndent;
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
        for (int i = 0; i < value.Length; i++) {
            if (value[i] != '\'') continue;
            WriteRun(value, start, i - start + 1);
            Write('\'');
            start = i + 1;
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
        char[] scratch = _scratch ??= new char[256];
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
        char[] scratch = _scratch ??= new char[256];
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
        if (_depth == _frames.Length) Array.Resize(ref _frames, _frames.Length * 2);
        _frames[_depth++] = frame;
    }

    private static InvalidOperationException Invalid(string message) => new InvalidOperationException(message);
}
