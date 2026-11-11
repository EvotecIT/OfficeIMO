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
    /// <summary>For an alias, the recording it expands to; for an anchored node, the recording its anchor names.</summary>
    public YamlRecording? Recording;
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

    /// <summary>
    /// The nodes an alias to this node stands for, counted as replaying it counts them: every node event, and for each
    /// alias inside, its own target's weight. Set when the recording closes.
    /// </summary>
    public long Weight { get; set; }

    /// <summary>The cursor's running weight when the node started, from which <see cref="Weight"/> is taken.</summary>
    public long WeightStart { get; set; }

    /// <summary>
    /// Whether the node's events are kept for replay. A collection whose value is shared among its aliases stops being
    /// kept (<see cref="YamlEventCursor.StopLogging"/>), and then nothing may replay it.
    /// </summary>
    public bool Logged { get; set; } = true;
}

/// <summary>
/// The event source the serializer reads: a <see cref="YamlReader"/> with aliases resolved, or the recorded events of a
/// <see cref="YamlNode"/> tree. Anchored nodes are recorded only while they are read, so input without anchors costs
/// nothing extra. An alias is expanded by replaying its anchored node's events; replayed nodes count against
/// <see cref="YamlReaderOptions.MaxAliasExpansion"/>. A reader that reuses the value read for the anchor instead charges
/// the same count through <see cref="ChargeAlias"/>.
/// </summary>
internal sealed class YamlEventCursor {
    private struct Frame {
        public List<YamlRecordedEvent> Log;
        public int Position;
        public int End;
        public bool CountsExpansion;
    }

    // A log longer than this is dropped instead of kept with a cursor returned for reuse.
    private const int MaxKeptLog = 4096;
    [ThreadStatic] private static YamlEventCursor? t_spare;

    private YamlReader? _reader;
    private bool _mergeKeysByDefault;
    private bool _strict;
    private int _maxExpansion;
    // Replay frames exist only for aliases, merges and recorded trees, so a live read allocates none.
    private Frame[] _frames = Array.Empty<Frame>();
    private int _frameCount;
    private YamlRecordedEvent _current;
    private bool _live;
    private long _expanded;
    private int _depth;
    // The weight of the node events recorded so far, from which each recording's weight is taken.
    private long _recordedWeight;
    // How many open anchored nodes keep their events; while none does, events are not logged.
    private int _logging;
    // The recording the current live event's anchor names; read only while the reader reports an anchor.
    private YamlRecording? _liveDefined;

    // Anchors of the current document, and the log their events are recorded in.
    private Dictionary<string, YamlRecording>? _anchors;
    private List<YamlRecordedEvent>? _log;
    private List<YamlRecording>? _open;
    private List<YamlRecording>? _captures;
    private int _currentLogIndex = -1;

    /// <summary>Creates a cursor over a reader; it starts on the reader's current event.</summary>
    public YamlEventCursor(YamlReader reader) => Attach(reader);

    /// <summary>
    /// Takes this thread's returned cursor, or creates one, over a reader. A stream of many small documents then reads
    /// each one without a cursor of its own; pair it with <see cref="Return"/> once the value is read.
    /// </summary>
    public static YamlEventCursor Rent(YamlReader reader) {
        YamlEventCursor? spare = t_spare;
        if (spare is null) return new YamlEventCursor(reader);
        t_spare = null;
        spare.Attach(reader);
        return spare;
    }

    /// <summary>Clears the cursor and keeps it for this thread's next read. The caller drops its reference.</summary>
    public void Return() {
        _reader = null;
        Array.Clear(_frames, 0, _frames.Length);
        _frameCount = 0;
        _current = default;
        _expanded = 0;
        _depth = 0;
        _recordedWeight = 0;
        _logging = 0;
        _liveDefined = null;
        _anchors?.Clear();
        if (_log is { Capacity: > MaxKeptLog }) _log = null;
        else _log?.Clear();
        _open?.Clear();
        _captures?.Clear();
        _currentLogIndex = -1;
        t_spare = this;
    }

    private void Attach(YamlReader reader) {
        _reader = reader;
        YamlReaderOptions options = reader.Options;
        _maxExpansion = options.MaxAliasExpansion;
        _strict = options.Strict;
        _mergeKeysByDefault = !options.Strict;
        _live = true;
    }

    /// <summary>Creates a cursor over recorded events, positioned on the first.</summary>
    public YamlEventCursor(List<YamlRecordedEvent> events, int maxExpansion, bool strict) {
        _maxExpansion = maxExpansion;
        _strict = strict;
        _log = events;
        PushFrame(events, 0, events.Count, countsExpansion: false);
        if (events.Count > 0) MoveReplay();
    }

    /// <summary>Gets whether the input is read strictly; a duplicate key then fails instead of replacing the earlier value.</summary>
    public bool Strict => _strict;

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
    /// <summary>Reads the current live event as a string when it is a scalar that is not null, as the string converter would.</summary>
    public bool TryReadLiveString(out string? text) {
        text = null;
        if (!_live) return false;
        YamlReader reader = _reader!;
        if (reader.TokenType != YamlTokenType.Scalar || reader.IsNullScalar()) return false;
        text = reader.GetString();
        return true;
    }

    /// <summary>
    /// Gets the current live event's scalar text and the schema to resolve it with, as <see cref="ScalarSpan"/> and
    /// <see cref="Schema"/> return them, when it is a scalar that is not null.
    /// </summary>
    public bool TryGetLiveScalar(out ReadOnlySpan<char> text, out YamlSchema schema) {
        text = default;
        schema = YamlSchema.Core;
        if (!_live) return false;
        YamlReader reader = _reader!;
        if (reader.TokenType != YamlTokenType.Scalar || reader.IsNullScalar()) return false;
        text = reader.ScalarValueSpan;
        YamlSchema documentSchema = reader.DocumentSchema;
        if (documentSchema != YamlSchema.Failsafe) schema = documentSchema;
        return true;
    }

    /// <summary>
    /// Gets the current live event's style, text and document schema when it is an untagged scalar, so an untyped read
    /// resolves it with one look at the reader instead of through <see cref="ResolveKind"/> and <see cref="ScalarSpan"/>.
    /// </summary>
    public bool TryGetLiveUntagged(out YamlStyle style, out ReadOnlySpan<char> text, out YamlSchema schema) {
        style = YamlStyle.None;
        text = default;
        schema = YamlSchema.Core;
        if (!_live) return false;
        YamlReader reader = _reader!;
        if (reader.TokenType != YamlTokenType.Scalar || reader.Tag is not null) return false;
        style = reader.Style;
        schema = reader.DocumentSchema;
        if (style == YamlStyle.Plain) text = reader.ScalarValueSpan;
        return true;
    }

    public bool IsNullScalar => _live ? _reader!.IsNullScalar() : _current.Type == YamlTokenType.Scalar && _current.IsNull;

    /// <summary>Gets whether the current scalar is the merge key <c>&lt;&lt;</c> where merge keys apply.</summary>
    public bool IsMergeKey => _live ? LiveIsMergeKey() : _current.IsMergeKey;

    /// <summary>
    /// Gets the recording the current node's anchor names when the node is read for the first time, or null: for a node
    /// without an anchor and for every node replayed by an alias, whose value is not the anchor's own.
    /// </summary>
    public YamlRecording? DefinedAnchor {
        get {
            if (_live) return _reader!.HasAnchor ? _liveDefined : null;
            return _current.Anchor is not null && _current.Type != YamlTokenType.Alias && !InExpansion() ? _current.Recording : null;
        }
    }

    /// <summary>Gets the recording the current alias refers to, or null when it refers to no anchor defined before it.</summary>
    public YamlRecording? AliasTarget {
        get {
            if (!_live) return _current.Type == YamlTokenType.Alias ? _current.Recording : null;
            return _anchors is not null && _anchors.TryGetValue(_reader!.GetString(), out YamlRecording? target) ? target : null;
        }
    }

    /// <summary>
    /// Counts the nodes an alias stands for against <see cref="YamlReaderOptions.MaxAliasExpansion"/> without replaying
    /// them, for a reader that reuses the value read for the anchor. A shared value stands for the same nodes wherever
    /// it is later walked, written or copied, so it costs what expanding it would.
    /// </summary>
    public void ChargeAlias(YamlRecording target) {
        _expanded += target.Weight;
        if (_expanded > _maxExpansion) throw new YamlLimitExceededException(nameof(YamlReaderOptions.MaxAliasExpansion), _maxExpansion, Position);
    }

    private bool InExpansion() {
        for (int i = 0; i < _frameCount; i++) {
            if (_frames[i].CountsExpansion) return true;
        }
        return false;
    }

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
        YamlReader reader = _reader;
        YamlTokenType type;
        while (true) {
            if (!reader.Read()) return false;
            type = reader.TokenType;
            if (type != YamlTokenType.Comment) break;
        }
        // Most events define no anchor and arrive while nothing records: only the depth changes.
        if (type != YamlTokenType.DocumentStart && !reader.HasAnchor && _open is not { Count: > 0 } && _captures is not { Count: > 0 }) {
            _currentLogIndex = -1;
            if (type is YamlTokenType.MappingStart or YamlTokenType.SequenceStart) _depth++;
            else if (type is YamlTokenType.MappingEnd or YamlTokenType.SequenceEnd) _depth--;
            return true;
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
            target = (_current.Type == YamlTokenType.Alias ? _current.Recording : null) ?? throw new YamlException(YamlErrorCodes.UndefinedAlias, "The alias '*" + _current.Value + "' refers to no anchor defined before it in the document.", Position);
        }
        if (target.End < 0) {
            throw new YamlException(YamlErrorCodes.RecursiveAlias, "The alias '*" + GetString() + "' is inside the node its anchor names; reading it would never end.", Position);
        }
        if (!target.Logged) {
            // Only a collection read as a shared untyped value stops being kept, and every untyped alias to it returns
            // that value; anything else that reads an alias to it reads a scalar.
            throw new YamlException(YamlErrorCodes.TypeMismatch, "The alias '*" + GetString() + "' names a mapping or sequence where a scalar is required.", Position);
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
                _logging = 0;
                break;
            case YamlTokenType.MappingEnd:
            case YamlTokenType.SequenceEnd:
                _depth--;
                break;
        }

        // Open anchored nodes need their weights and ends; only the logged ones need their events.
        bool open = _open is { Count: > 0 };
        bool node = type is YamlTokenType.Scalar or YamlTokenType.Alias or YamlTokenType.MappingStart or YamlTokenType.SequenceStart;
        YamlRecording? defined = null;
        if (node && reader.HasAnchor) {
            _log ??= new List<YamlRecordedEvent>();
            defined = new YamlRecording(_log, _log.Count) { Depth = _depth, WeightStart = _recordedWeight };
            (_open ??= new List<YamlRecording>()).Add(defined);
            (_anchors ??= new Dictionary<string, YamlRecording>(StringComparer.Ordinal))[reader.Anchor!] = defined;
            _liveDefined = defined;
            _logging++;
            open = true;
        }
        bool logging = _logging > 0;
        if (logging) {
            _currentLogIndex = _log!.Count;
            YamlRecordedEvent recorded = Snapshot();
            if (defined is not null) recorded.Recording = defined;
            if (node) _recordedWeight += type == YamlTokenType.Alias ? 1 + (recorded.Recording?.Weight ?? 0) : 1;
            _log.Add(recorded);
        } else if (open && node) {
            _recordedWeight += type == YamlTokenType.Alias && _anchors!.TryGetValue(reader.GetString(), out YamlRecording? target) ? 1 + target.Weight : 1;
        }
        if (_captures is { Count: > 0 }) {
            YamlRecordedEvent snapshot = logging ? _log![_currentLogIndex] : Snapshot();
            // A captured node is read for the first time when the capture is replayed, so it keeps its anchor.
            if (node && reader.HasAnchor) snapshot.Anchor = reader.Anchor;
            foreach (YamlRecording capture in _captures) capture.Log.Add(snapshot);
        }

        if (type is YamlTokenType.MappingStart or YamlTokenType.SequenceStart) _depth++;
        if (open) CloseRecordings();
    }

    /// <summary>
    /// Stops keeping the events of the anchored collection that starts at the current live event, for a reader that
    /// shares its value among every alias to it and so never replays it. Its weight and end are still tracked.
    /// </summary>
    public void StopLogging(YamlRecording anchored) {
        if (!_live || !anchored.Logged || anchored.End >= 0) return;
        anchored.Logged = false;
        // The collection's start is the last logged event; drop it when no enclosing node or capture keeps it.
        if (--_logging == 0 && _currentLogIndex == _log!.Count - 1 && _captures is not { Count: > 0 }) {
            _log.RemoveAt(_currentLogIndex);
            _currentLogIndex = -1;
        }
    }

    /// <summary>Closes the recordings whose node ended with the current event.</summary>
    private void CloseRecordings() {
        List<YamlRecording> open = _open!;
        for (int i = open.Count - 1; i >= 0; i--) {
            YamlRecording recording = open[i];
            if (_depth > recording.Depth) break;
            recording.End = _log!.Count;
            recording.Weight = _recordedWeight - recording.WeightStart;
            if (recording.Logged) _logging--;
            open.RemoveAt(i);
        }
    }

    /// <summary>Copies the current event so it can be replayed after the reader moves on.</summary>
    private YamlRecordedEvent Snapshot() {
        if (!_live) {
            YamlRecordedEvent copy = _current;
            copy.Anchor = null;
            if (copy.Type != YamlTokenType.Alias) copy.Recording = null;
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
                if (_anchors is not null && _anchors.TryGetValue(name, out YamlRecording? target)) snapshot.Recording = target;
                break;
        }
        return snapshot;
    }

    private bool LiveIsMergeKey() {
        YamlReader reader = _reader!;
        if (reader.TokenType != YamlTokenType.Scalar) return false;
        // Most keys are untagged text of another shape, which needs neither the tag nor the decoded text.
        if (reader.IsUntaggedNonMergeText) return false;
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
        long weight = 0;
        RecordNode(node, log, anchors, ref weight);
        return log;
    }

    private static void RecordNode(YamlNode node, List<YamlRecordedEvent> log, Dictionary<YamlNode, YamlRecording> anchors, ref long weight) {
        YamlRecording? recording = null;
        if (node.Anchor is not null && node is not YamlAlias) {
            recording = new YamlRecording(log, log.Count) { WeightStart = weight };
            anchors[node] = recording;
        }
        weight++;
        var ev = new YamlRecordedEvent { Style = node.Style, Tag = node.Tag, Anchor = node.Anchor, Node = node.HasSourceSpan ? node : null, Schema = node.Schema, Recording = recording };
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
                    long ownWeight = 0;
                    RecordNode(alias.Target, own, anchors, ref ownWeight);
                    target = anchors[alias.Target];
                }
                ev.Recording = target;
                weight += target.Weight;
                log.Add(ev);
                break;
            case YamlMapping mapping:
                ev.Type = YamlTokenType.MappingStart;
                log.Add(ev);
                foreach (YamlMappingEntry entry in mapping) {
                    RecordNode(entry.Key, log, anchors, ref weight);
                    RecordNode(entry.Value, log, anchors, ref weight);
                }
                log.Add(new YamlRecordedEvent { Type = YamlTokenType.MappingEnd, Schema = ev.Schema });
                break;
            case YamlSequence sequence:
                ev.Type = YamlTokenType.SequenceStart;
                log.Add(ev);
                foreach (YamlNode item in sequence) RecordNode(item, log, anchors, ref weight);
                log.Add(new YamlRecordedEvent { Type = YamlTokenType.SequenceEnd, Schema = ev.Schema });
                break;
        }
        if (recording is not null) {
            recording.End = log.Count;
            recording.Weight = weight - recording.WeightStart;
        }
    }
}
