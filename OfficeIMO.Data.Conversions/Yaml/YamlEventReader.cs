using OfficeIMO.Yaml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Streams a YAML document stream as data events through <see cref="YamlReader"/>. Aliases replay their anchored node
/// (bounded by <see cref="YamlReaderOptions.MaxAliasExpansion"/>), and merge keys resolve under the YAML 1.1 schema or
/// an explicit <c>!!merge</c> tag. Only anchored nodes and mappings with merge keys are buffered.
/// </summary>
internal sealed class YamlEventReader : DataEventReader {
    private readonly YamlReader _reader;
    private bool _yaml11;
    private bool _mergeKeys;
    private readonly bool _strict;
    private int _forwardedFindings;
    private readonly int _maxAliasExpansion;
    private readonly Stack<Frame> _frames = new();
    private readonly Dictionary<string, DataNode> _anchors = new(StringComparer.Ordinal);
    private readonly List<Recording> _recordings = new();
    private readonly Stack<Replay> _replays = new();
    private long _expanded;

    private struct Frame {
        public bool IsMapping;
        public bool ExpectKey;
    }

    private sealed class Recording {
        public Recording(string name) {
            Name = name;
        }

        public string Name { get; }
        public DataNodeBuilder Builder { get; } = new DataNodeBuilder();
    }

    private sealed class Replay {
        public Replay(DataNodeReplay events, bool closesMapping) {
            Events = events;
            ClosesMapping = closesMapping;
        }

        public DataNodeReplay Events { get; }
        public bool ClosesMapping { get; }
    }

    public YamlEventReader(Stream stream, DataConversionContext context) : base(context) {
        YamlReaderOptions? source = context.Options.YamlReader;
        var options = new YamlReaderOptions { IncludeComments = true };
        if (source != null) {
            options.Schema = source.Schema;
            options.MaxDepth = source.MaxDepth;
            options.MaxNodes = source.MaxNodes;
            options.MaxAliasExpansion = source.MaxAliasExpansion;
            options.MaxScalarLength = source.MaxScalarLength;
            options.MaxInputBytes = source.MaxInputBytes;
            options.Strict = source.Strict;
        }

        _strict = options.Strict;
        _maxAliasExpansion = options.MaxAliasExpansion;
        _reader = new YamlReader(stream, options, context.CancellationToken);
        TrackDuplicateKeys = true;
    }

    public override string? Location => _reader.Start.Line > 0
        ? "line " + _reader.Start.Line.ToString(CultureInfo.InvariantCulture) + ", column " + _reader.Start.Column.ToString(CultureInfo.InvariantCulture)
        : null;

    private bool InKeyPosition => _frames.Count > 0 && _frames.Peek().IsMapping && _frames.Peek().ExpectKey;

    protected override bool ReadCore() {
        while (true) {
            if (_replays.Count > 0) {
                Replay replay = _replays.Peek();
                if (replay.Events.Next(out DataTokenType type, out DataScalarKind kind, out string? text)) {
                    Emit(type, kind, text.AsSpan());
                    return true;
                }

                _replays.Pop();
                if (replay.ClosesMapping) {
                    _frames.Pop();
                    Emit(DataTokenType.EndObject, default, default);
                    AfterValue();
                    return true;
                }

                AfterValue();
                continue;
            }

            if (!_reader.Read()) {
                ForwardFindings();
                return false;
            }

            switch (_reader.TokenType) {
                case YamlTokenType.DocumentStart:
                    // Each document resolves with its own schema, as YamlDocument does: %YAML 1.1 selects 1.1.
                    _yaml11 = _reader.Schema == YamlSchema.Yaml11;
                    _mergeKeys = _yaml11 || (_reader.Schema == YamlSchema.Core && !_strict);
                    _anchors.Clear();
                    Emit(DataTokenType.StartDocument, default, default);
                    return true;
                case YamlTokenType.DocumentEnd:
                    Emit(DataTokenType.EndDocument, default, default);
                    return true;
                case YamlTokenType.Comment:
                    SetComment(_reader.ValueSpan);
                    return true;
                case YamlTokenType.MappingStart:
                case YamlTokenType.SequenceStart:
                    if (InKeyPosition) {
                        EmitComplexKey(Capture());
                        return true;
                    }

                    YamlScalarClassifier.CheckCollectionTag(_reader.Tag, Context);
                    bool mapping = _reader.TokenType == YamlTokenType.MappingStart;
                    if (_reader.Anchor is { } anchor) _recordings.Add(new Recording(anchor));
                    _frames.Push(new Frame { IsMapping = mapping, ExpectKey = true });
                    Emit(mapping ? DataTokenType.StartObject : DataTokenType.StartArray, default, default);
                    return true;
                case YamlTokenType.MappingEnd:
                case YamlTokenType.SequenceEnd:
                    _frames.Pop();
                    Emit(_reader.TokenType == YamlTokenType.MappingEnd ? DataTokenType.EndObject : DataTokenType.EndArray, default, default);
                    AfterValue();
                    return true;
                case YamlTokenType.Scalar:
                    if (InKeyPosition) {
                        var key = new YamlReaderScalar(_reader);
                        if (IsMergeKey(ref key)) {
                            if (StartMerge()) return true;
                            continue;
                        }

                        if (_reader.Anchor is { } keyAnchor) {
                            var anchored = new YamlReaderScalar(_reader);
                            DataScalarKind keyKind = YamlScalarClassifier.Classify(ref anchored, _yaml11, Context, out ReadOnlySpan<char> keyValue);
                            _anchors[keyAnchor] = DataNode.Scalar(keyKind, keyValue.ToString());
                        }

                        ReadOnlySpan<char> name = YamlScalarClassifier.KeyText(ref key, _yaml11, Context);
                        Emit(DataTokenType.PropertyName, default, name);
                        SetExpectKey(false);
                        return true;
                    }

                    var scalar = new YamlReaderScalar(_reader);
                    DataScalarKind scalarKind = YamlScalarClassifier.Classify(ref scalar, _yaml11, Context, out ReadOnlySpan<char> value);
                    if (_reader.Anchor is { } scalarAnchor) _anchors[scalarAnchor] = DataNode.Scalar(scalarKind, value.ToString());
                    Emit(DataTokenType.Scalar, scalarKind, value);
                    AfterValue();
                    return true;
                case YamlTokenType.Alias:
                    DataNode target = Resolve(_reader.GetString());
                    if (InKeyPosition) {
                        if (target.Type == DataTokenType.Scalar) {
                            Emit(DataTokenType.PropertyName, default, (target.Kind == DataScalarKind.Null ? "null" : target.Text).AsSpan());
                            SetExpectKey(false);
                        } else {
                            EmitComplexKey(target);
                        }

                        return true;
                    }

                    _replays.Push(new Replay(new DataNodeReplay(target), closesMapping: false));
                    continue;
            }
        }
    }

    private void SetExpectKey(bool value) {
        Frame frame = _frames.Pop();
        frame.ExpectKey = value;
        _frames.Push(frame);
    }

    private void AfterValue() {
        if (_frames.Count > 0 && _frames.Peek().IsMapping) SetExpectKey(true);
    }

    private void EmitComplexKey(DataNode key) {
        Context.Report(OfficeDiagnosticCodes.YamlComplexKey, OfficeConversionLossKind.Approximation,
            "Mappings or sequences used as mapping keys were written as their flow YAML text.");
        Emit(DataTokenType.PropertyName, default, key.ToFlowText().AsSpan());
        SetExpectKey(false);
    }

    private void Emit(DataTokenType type, DataScalarKind kind, ReadOnlySpan<char> text) {
        switch (type) {
            case DataTokenType.PropertyName: SetPropertyName(text); break;
            case DataTokenType.Scalar: SetScalar(kind, text); break;
            default: Set(type); break;
        }

        if (_recordings.Count == 0 || type is DataTokenType.StartDocument or DataTokenType.EndDocument) return;
        for (int i = _recordings.Count - 1; i >= 0; i--) {
            Recording recording = _recordings[i];
            recording.Builder.Add(type, kind, text);
            if (recording.Builder.IsComplete) {
                _anchors[recording.Name] = recording.Builder.Result!;
                _recordings.RemoveAt(i);
            }
        }
    }

    private DataNode Resolve(string name) {
        if (!_anchors.TryGetValue(name, out DataNode? node)) {
            throw new InvalidDataException("The alias *" + name + " refers to an anchor that is not defined before it" + At() + ".");
        }

        _expanded += node.Size;
        if (_expanded > _maxAliasExpansion) {
            throw new InvalidDataException("Expanding aliases visited more than " + _maxAliasExpansion.ToString(CultureInfo.InvariantCulture)
                + " nodes (MaxAliasExpansion)" + At() + ".");
        }

        Context.Report(OfficeDiagnosticCodes.YamlAliasExpanded, OfficeConversionLossKind.None,
            "YAML aliases were expanded into copies of their anchored values; the data is unchanged.");
        return node;
    }

    private bool IsMergeKey(ref YamlReaderScalar key) {
        string? tag = key.Tag;
        if (tag != null) return tag == YamlScalarClassifier.TagPrefix + "merge";
        return _mergeKeys && key.Style == YamlStyle.Plain && key.Text.SequenceEqual("<<".AsSpan());
    }

    // The engine's own findings (lenient indentation) belong in the conversion report too.
    private void ForwardFindings() {
        IReadOnlyList<OfficeDiagnostic> findings = _reader.Diagnostics;
        for (; _forwardedFindings < findings.Count; _forwardedFindings++) {
            OfficeDiagnostic finding = findings[_forwardedFindings];
            Context.Report(finding.Code, finding.LossKind, finding.Message, finding.Location);
        }
    }

    /// <summary>
    /// Resolves a merge key: the rest of the mapping is buffered so explicit keys can override merged ones, then merged
    /// keys not set explicitly are replayed, followed by the explicit keys.
    /// </summary>
    private bool StartMerge() {
        if (!_reader.Read()) throw new InvalidDataException("A merge key has no value.");
        var sources = new List<DataNode>();
        DataNode value = Capture();
        if (value.Type == DataTokenType.StartArray) sources.AddRange(value.Items!);
        else sources.Add(value);
        DataNode rest = CaptureMappingRest();
        var members = new List<KeyValuePair<string, DataNode>>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, DataNode> property in rest.Properties!) seen.Add(property.Key);
        foreach (DataNode source in sources) {
            if (source.Type != DataTokenType.StartObject) throw new InvalidDataException("A merge key must refer to mappings" + At() + ".");
            foreach (KeyValuePair<string, DataNode> property in source.Properties!) {
                if (seen.Add(property.Key) && !HasKey(property.Key.AsSpan())) members.Add(property);
            }
        }

        members.AddRange(rest.Properties!);
        _replays.Push(new Replay(DataNodeReplay.Members(members), closesMapping: true));
        return false;
    }

    /// <summary>Reads the node at the current event into a buffer, resolving aliases and nested merge keys.</summary>
    private DataNode Capture() {
        string? anchor = _reader.Anchor;
        DataNode node;
        switch (_reader.TokenType) {
            case YamlTokenType.Scalar: {
                var scalar = new YamlReaderScalar(_reader);
                DataScalarKind kind = YamlScalarClassifier.Classify(ref scalar, _yaml11, Context, out ReadOnlySpan<char> text);
                node = DataNode.Scalar(kind, text.ToString());
                break;
            }
            case YamlTokenType.Alias:
                return Resolve(_reader.GetString());
            case YamlTokenType.SequenceStart:
                YamlScalarClassifier.CheckCollectionTag(_reader.Tag, Context);
                node = DataNode.Array();
                while (Next() != YamlTokenType.SequenceEnd) node.Items!.Add(Capture());
                break;
            case YamlTokenType.MappingStart:
                YamlScalarClassifier.CheckCollectionTag(_reader.Tag, Context);
                node = CaptureMappingRest();
                break;
            default:
                throw new InvalidDataException("Unexpected YAML event " + _reader.TokenType + At() + ".");
        }

        if (anchor != null) _anchors[anchor] = node;
        return node;
    }

    /// <summary>Reads key/value pairs up to the end of the current mapping, applying merge keys.</summary>
    private DataNode CaptureMappingRest() {
        var explicitMembers = new List<KeyValuePair<string, DataNode>>();
        var merged = new List<DataNode>();
        while (Next() != YamlTokenType.MappingEnd) {
            string key;
            if (_reader.TokenType == YamlTokenType.Scalar) {
                var scalar = new YamlReaderScalar(_reader);
                if (IsMergeKey(ref scalar)) {
                    Next();
                    DataNode source = Capture();
                    if (source.Type == DataTokenType.StartArray) merged.AddRange(source.Items!);
                    else merged.Add(source);
                    continue;
                }

                key = YamlScalarClassifier.KeyText(ref scalar, _yaml11, Context).ToString();
            } else {
                DataNode complex = Capture();
                key = complex.Type == DataTokenType.Scalar ? complex.Text ?? "null" : complex.ToFlowText();
                if (complex.Type != DataTokenType.Scalar) {
                    Context.Report(OfficeDiagnosticCodes.YamlComplexKey, OfficeConversionLossKind.Approximation,
                        "Mappings or sequences used as mapping keys were written as their flow YAML text.");
                }
            }

            foreach (KeyValuePair<string, DataNode> existing in explicitMembers) {
                if (existing.Key == key) throw new InvalidDataException("The key '" + key + "' appears twice in one object" + At() + ".");
            }

            Next();
            explicitMembers.Add(new KeyValuePair<string, DataNode>(key, Capture()));
        }

        DataNode mapping = DataNode.Object();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, DataNode> property in explicitMembers) seen.Add(property.Key);
        foreach (DataNode source in merged) {
            if (source.Type != DataTokenType.StartObject) throw new InvalidDataException("A merge key must refer to mappings" + At() + ".");
            foreach (KeyValuePair<string, DataNode> property in source.Properties!) {
                if (seen.Add(property.Key)) mapping.Properties!.Add(property);
            }
        }

        mapping.Properties!.AddRange(explicitMembers);
        return mapping;
    }

    // Advances past comments, which a buffered node does not keep.
    private YamlTokenType Next() {
        do {
            if (!_reader.Read()) throw new InvalidDataException("The YAML input ended inside a node.");
        } while (_reader.TokenType == YamlTokenType.Comment);

        return _reader.TokenType;
    }

    public override void Dispose() => _reader.Dispose();
}
