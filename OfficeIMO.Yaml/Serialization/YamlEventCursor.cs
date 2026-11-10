namespace OfficeIMO.Yaml;

/// <summary>A node event kept for replay: the events of an anchored node, a captured mapping or a model subtree.</summary>
internal struct YamlRecordedEvent {
    public YamlTokenType Type;
    public YamlStyle Style;
    public bool IsNull;
    public bool IsMergeKey;
    public YamlSchema Schema;
    public string? Tag;
    /// <summary>The decoded scalar value, or the alias name.</summary>
    public string? Value;
    /// <summary>The anchor exposed on replay: set for model subtrees, whose anchors are first seen there.</summary>
    public string? Anchor;
    public YamlPosition Position;
    /// <summary>The model node the event came from, whose position is computed only when an error needs it.</summary>
    public YamlNode? Node;
    /// <summary>For an alias, the recording it expands to.</summary>
    public YamlRecording? Target;
}

/// <summary>A range of recorded events: one node, or the events captured while scanning a mapping.</summary>
internal sealed class YamlRecording {
    public YamlRecording(List<YamlRecordedEvent> log, int start) {
        Log = log;
        Start = start;
        End = -1;
    }

    public List<YamlRecordedEvent> Log { get; }

    public int Start { get; }

    /// <summary>The index after the node's last event, or -1 while the node is still being read.</summary>
    public int End { get; set; }

    /// <summary>The live nesting depth before the node's first event, which its last event returns to.</summary>
    public int Depth { get; set; }
}

/// <summary>
/// The event source the serializer reads: a <see cref="YamlReader"/> with aliases resolved, or the recorded events of a
/// <see cref="YamlNode"/> tree. Anchored nodes are recorded only while they are read, so input without anchors costs
/// nothing extra. An alias is expanded by replaying its anchored node's events; replayed nodes count against
/// <see cref="YamlReaderOptions.MaxAliasExpansion"/>.
/// </summary>
internal sealed class YamlEventCursor {
    private struct Frame {
        public List<YamlRecordedEvent> Log;
        public int Position;
        public int End;
        public bool CountsExpansion;
    }

    private readonly YamlReader? _reader;
    private readonly bool _mergeKeysByDefault;
    private readonly bool _strict;
    private readonly int _maxExpansion;
    // Replay frames exist only for aliases, merges and recorded trees, so a live read allocates none.
    private Frame[] _frames = Array.Empty<Frame>();
    private int _frameCount;
    private YamlRecordedEvent _current;
    private bool _live;
    private long _expanded;
    private int _depth;

    // Anchors of the current document, and the log their events are recorded in.
    private Dictionary<string, YamlRecording>? _anchors;
    private List<YamlRecordedEvent>? _log;
    private List<YamlRecording>? _open;
    private List<YamlRecording>? _captures;
    private int _currentLogIndex = -1;

    /// <summary>Creates a cursor over a reader; it starts on the reader's current event.</summary>
    public YamlEventCursor(YamlReader reader) {
        _reader = reader;
        YamlReaderOptions options = reader.Options;
        _maxExpansion = options.MaxAliasExpansion;
        _strict = options.Strict;
        _mergeKeysByDefault = !options.Strict;
        _live = true;
    }

    /// <summary>Creates a cursor over recorded events, positioned on the first.</summary>
    public YamlEventCursor(List<YamlRecordedEvent> events, int maxExpansion) {
        _maxExpansion = maxExpansion;
        _log = events;
        PushFrame(events, 0, events.Count, countsExpansion: false);
        if (events.Count > 0) MoveReplay();
    }

    public YamlTokenType TokenType => _live ? _reader!.TokenType : _current.Type;

    public YamlStyle Style => _live ? _reader!.Style : _current.Style;

    /// <summary>Gets the resolved tag of the current node, or null.</summary>
    public string? Tag => _live ? _reader!.Tag : _current.Tag;

    /// <summary>Gets the anchor defined on the current node, or null. Replayed copies report no anchor.</summary>
    public string? Anchor => _live ? (_reader!.HasAnchor ? _reader.Anchor : null) : _current.Anchor;

    public YamlPosition Position {
        get {
            if (_live) return _reader!.Start;
            return _current.Node is YamlNode node ? node.Start : _current.Position;
        }
    }

    /// <summary>Gets the schema numbers and booleans parse with in the current document (core for failsafe).</summary>
    public YamlSchema Schema {
        get {
            YamlSchema schema = _live ? _reader!.DocumentSchema : _current.Schema;
            return schema == YamlSchema.Failsafe ? YamlSchema.Core : schema;
        }
    }

    /// <summary>Resolves the current scalar as the schema does: tags first, then style, then the plain text.</summary>
    public YamlScalarKind ResolveKind() {
        YamlSchema schema = _live ? _reader!.DocumentSchema : _current.Schema;
        YamlStyle style = Style;
        bool plain = Tag is null && style == YamlStyle.Plain;
        if (!YamlSchemaResolver.TryGetKind(Tag, style, plain ? ScalarSpan : default, schema, out YamlScalarKind kind)) return YamlScalarKind.Custom;
        return kind;
    }

    /// <summary>Gets whether the current node is a null scalar under its schema: plain <c>null</c>, <c>~</c>, empty, or <c>!!null</c>.</summary>
    public bool IsNullScalar => _live ? _reader!.IsNullScalar() : _current.Type == YamlTokenType.Scalar && _current.IsNull;

    /// <summary>Gets whether the current scalar is the merge key <c>&lt;&lt;</c> where merge keys apply.</summary>
    public bool IsMergeKey => _live ? LiveIsMergeKey() : _current.IsMergeKey;

    /// <summary>Gets the decoded value of the current scalar, or the name of the current alias.</summary>
    public string GetString() => _live ? _reader!.GetString() : _current.Value ?? string.Empty;

    /// <summary>Gets the current scalar's value without allocating when the source already holds it verbatim.</summary>
    public ReadOnlySpan<char> ScalarSpan => _live ? _reader!.ScalarValueSpan : (_current.Value ?? string.Empty).AsSpan();

    /// <summary>Gets whether the cursor reads events recorded from a model tree, where positions are optional.</summary>
    public bool HasReader => _reader is not null;

    /// <summary>Takes the reader's current event as the first one, recording it when it defines an anchor.</summary>
    public void Begin() {
        if (_reader is not null && _live && _reader.TokenType != YamlTokenType.None) OnLiveEvent();
    }

    /// <summary>Advances to the next event, skipping comments. Returns false after the last event.</summary>
    public bool Read() {
        while (_frameCount > 0) {
            ref Frame frame = ref _frames[_frameCount - 1];
            if (frame.Position < frame.End) {
                MoveReplay();
                return true;
            }
            _frameCount--;
        }
        if (_reader is null) {
            _current = default;
            return false;
        }
        _live = true;
        while (true) {
            if (!_reader.Read()) return false;
            if (_reader.TokenType != YamlTokenType.Comment) break;
        }
        OnLiveEvent();
        return true;
    }

    /// <summary>Skips the current node, leaving the cursor on its last event. Aliases are not expanded.</summary>
    public void Skip() {
        YamlTokenType type = TokenType;
        if (type is not (YamlTokenType.MappingStart or YamlTokenType.SequenceStart)) return;
        int depth = 1;
        while (depth > 0) {
            if (!Read()) throw YamlSerializationErrors.UnexpectedEnd(this);
            switch (TokenType) {
                case YamlTokenType.MappingStart:
                case YamlTokenType.SequenceStart:
                    depth++;
                    break;
                case YamlTokenType.MappingEnd:
                case YamlTokenType.SequenceEnd:
                    depth--;
                    break;
            }
        }
    }

    /// <summary>
    /// Replaces the current alias with the events of its anchored node and moves to the first of them. The node's
    /// events count against the alias-expansion limit.
    /// </summary>
    public void ExpandAlias() {
        YamlRecording? target;
        if (_live) {
            string name = _reader!.GetString();
            if (_anchors is null || !_anchors.TryGetValue(name, out target)) {
                throw new YamlException(YamlErrorCodes.UndefinedAlias, "The alias '*" + name + "' refers to no anchor defined before it in the document.", Position);
            }
        } else {
            target = _current.Target ?? throw new YamlException(YamlErrorCodes.UndefinedAlias, "The alias '*" + _current.Value + "' refers to no anchor defined before it in the document.", Position);
        }
        if (target.End < 0) {
            throw new YamlException(YamlErrorCodes.RecursiveAlias, "The alias '*" + GetString() + "' is inside the node its anchor names; reading it would never end.", Position);
        }
        PushFrame(target.Log, target.Start, target.End, countsExpansion: true);
        MoveReplay();
    }

    /// <summary>
    /// Starts recording the events from the current one, for a mapping that must be scanned before it can be read.
    /// <see cref="EndCapture"/> closes the recording; <see cref="Replay"/> reads it again.
    /// </summary>
    public YamlRecording BeginCapture() {
        var capture = new YamlRecording(new List<YamlRecordedEvent>(), 0);
        capture.Log.Add(Snapshot());
        (_captures ??= new List<YamlRecording>()).Add(capture);
        return capture;
    }

    /// <summary>Stops recording into <paramref name="capture"/>.</summary>
    public void EndCapture(YamlRecording capture) {
        capture.End = capture.Log.Count;
        _captures!.Remove(capture);
    }

    /// <summary>Moves to the first event of a closed capture. Replaying a capture does not count as alias expansion.</summary>
    public void Replay(YamlRecording capture) {
        PushFrame(capture.Log, capture.Start, capture.End, countsExpansion: false);
        MoveReplay();
    }

    private void PushFrame(List<YamlRecordedEvent> log, int start, int end, bool countsExpansion) {
        if (_frameCount == _frames.Length) Array.Resize(ref _frames, Math.Max(4, _frames.Length * 2));
        _frames[_frameCount++] = new Frame { Log = log, Position = start, End = end, CountsExpansion = countsExpansion };
    }

    private void MoveReplay() {
        ref Frame frame = ref _frames[_frameCount - 1];
        _current = frame.Log[frame.Position++];
        _live = false;
        if (frame.CountsExpansion && _current.Type is YamlTokenType.Scalar or YamlTokenType.Alias or YamlTokenType.MappingStart or YamlTokenType.SequenceStart) {
            if (++_expanded > _maxExpansion) {
                throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxAliasExpansion), _maxExpansion, Position);
            }
        }
        if (_captures is { Count: > 0 }) {
            foreach (YamlRecording capture in _captures) capture.Log.Add(_current);
        }
    }

    // ------------------------------------------------------------------ live events and anchors

    private void OnLiveEvent() {
        YamlReader reader = _reader!;
        YamlTokenType type = reader.TokenType;
        _currentLogIndex = -1;
        switch (type) {
            case YamlTokenType.DocumentStart:
                // Anchors are scoped to their document.
                _anchors?.Clear();
                _log?.Clear();
                _open?.Clear();
                break;
            case YamlTokenType.MappingEnd:
            case YamlTokenType.SequenceEnd:
                _depth--;
                break;
        }

        bool recording = _open is { Count: > 0 };
        bool node = type is YamlTokenType.Scalar or YamlTokenType.Alias or YamlTokenType.MappingStart or YamlTokenType.SequenceStart;
        if (node && reader.HasAnchor) {
            _log ??= new List<YamlRecordedEvent>();
            var recordingOfNode = new YamlRecording(_log, _log.Count) { Depth = _depth };
            (_open ??= new List<YamlRecording>()).Add(recordingOfNode);
            (_anchors ??= new Dictionary<string, YamlRecording>(StringComparer.Ordinal))[reader.Anchor!] = recordingOfNode;
            recording = true;
        }
        if (recording) {
            _currentLogIndex = _log!.Count;
            _log.Add(Snapshot());
        }
        if (_captures is { Count: > 0 }) {
            YamlRecordedEvent snapshot = recording ? _log![_currentLogIndex] : Snapshot();
            // A captured node is read for the first time when the capture is replayed, so it keeps its anchor.
            if (node && reader.HasAnchor) snapshot.Anchor = reader.Anchor;
            foreach (YamlRecording capture in _captures) capture.Log.Add(snapshot);
        }

        if (type is YamlTokenType.MappingStart or YamlTokenType.SequenceStart) _depth++;
        if (recording) CloseRecordings();
    }

    /// <summary>Closes the recordings whose node ended with the current event.</summary>
    private void CloseRecordings() {
        List<YamlRecording> open = _open!;
        for (int i = open.Count - 1; i >= 0; i--) {
            YamlRecording recording = open[i];
            if (_depth > recording.Depth) break;
            recording.End = _log!.Count;
            open.RemoveAt(i);
        }
    }

    /// <summary>Copies the current event so it can be replayed after the reader moves on.</summary>
    private YamlRecordedEvent Snapshot() {
        if (!_live) {
            YamlRecordedEvent copy = _current;
            copy.Anchor = null;
            return copy;
        }
        YamlReader reader = _reader!;
        var snapshot = new YamlRecordedEvent {
            Type = reader.TokenType,
            Style = reader.Style,
            Schema = reader.DocumentSchema,
            Position = reader.Start
        };
        switch (snapshot.Type) {
            case YamlTokenType.Scalar:
                snapshot.Tag = reader.Tag;
                snapshot.Value = reader.GetString();
                snapshot.IsNull = reader.IsNullScalar();
                snapshot.IsMergeKey = LiveIsMergeKey();
                break;
            case YamlTokenType.MappingStart:
            case YamlTokenType.SequenceStart:
                snapshot.Tag = reader.Tag;
                break;
            case YamlTokenType.Alias:
                string name = reader.GetString();
                snapshot.Value = name;
                if (_anchors is not null && _anchors.TryGetValue(name, out YamlRecording? target)) snapshot.Target = target;
                break;
        }
        return snapshot;
    }

    private bool LiveIsMergeKey() {
        YamlReader reader = _reader!;
        if (reader.TokenType != YamlTokenType.Scalar) return false;
        string? tag = reader.Tag;
        if (tag is not null) return tag == YamlSchemaResolver.TagPrefix + "merge";
        if (reader.Style != YamlStyle.Plain) return false;
        ReadOnlySpan<char> text = reader.ScalarValueSpan;
        if (text.Length != 2 || text[0] != '<' || text[1] != '<') return false;
        YamlSchema schema = reader.DocumentSchema;
        return schema == YamlSchema.Yaml11 || (schema == YamlSchema.Core && _mergeKeysByDefault);
    }

    // ------------------------------------------------------------------ model trees

    /// <summary>Records the events of a model subtree, with anchors exposed and aliases bound to their anchored nodes.</summary>
    public static List<YamlRecordedEvent> Record(YamlNode node) {
        var log = new List<YamlRecordedEvent>();
        var anchors = new Dictionary<YamlNode, YamlRecording>();
        RecordNode(node, log, anchors);
        return log;
    }

    private static void RecordNode(YamlNode node, List<YamlRecordedEvent> log, Dictionary<YamlNode, YamlRecording> anchors) {
        YamlRecording? recording = null;
        if (node.Anchor is not null && node is not YamlAlias) {
            recording = new YamlRecording(log, log.Count);
            anchors[node] = recording;
        }
        var ev = new YamlRecordedEvent { Style = node.Style, Tag = node.Tag, Anchor = node.Anchor, Node = node.HasSourceSpan ? node : null, Schema = node.Schema };
        switch (node) {
            case YamlScalar scalar:
                ev.Type = YamlTokenType.Scalar;
                ev.Value = scalar.Value;
                ev.IsNull = scalar.ResolvesToNull();
                ev.IsMergeKey = scalar.IsMergeKey();
                if (ev.Style == YamlStyle.None) ev.Style = YamlStyle.DoubleQuoted;
                log.Add(ev);
                break;
            case YamlAlias alias:
                ev.Type = YamlTokenType.Alias;
                ev.Anchor = null;
                ev.Tag = null;
                ev.Value = alias.Name;
                if (!anchors.TryGetValue(alias.Target, out YamlRecording? target)) {
                    // An alias written before its anchor by editing: record the anchored node on its own.
                    var own = new List<YamlRecordedEvent>();
                    RecordNode(alias.Target, own, anchors);
                    target = anchors[alias.Target];
                }
                ev.Target = target;
                log.Add(ev);
                break;
            case YamlMapping mapping:
                ev.Type = YamlTokenType.MappingStart;
                log.Add(ev);
                foreach (YamlMappingEntry entry in mapping) {
                    RecordNode(entry.Key, log, anchors);
                    RecordNode(entry.Value, log, anchors);
                }
                log.Add(new YamlRecordedEvent { Type = YamlTokenType.MappingEnd, Schema = ev.Schema });
                break;
            case YamlSequence sequence:
                ev.Type = YamlTokenType.SequenceStart;
                log.Add(ev);
                foreach (YamlNode item in sequence) RecordNode(item, log, anchors);
                log.Add(new YamlRecordedEvent { Type = YamlTokenType.SequenceEnd, Schema = ev.Schema });
                break;
        }
        if (recording is not null) recording.End = log.Count;
    }
}
