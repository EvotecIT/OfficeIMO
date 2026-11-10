namespace OfficeIMO.Yaml;

/// <summary>
/// Builds the model from <see cref="YamlReader"/> events in one pass without recursion. Children of open collections
/// wait on one shared stack and move into an exactly sized array when their collection closes, so a collection costs
/// its node object and one array. The builder also enforces the per-document node and alias-expansion limits, resolves
/// aliases and checks mappings for duplicate keys.
/// </summary>
internal sealed class YamlTreeBuilder {
    /// <summary>Saturation point for expansion sizes, far above any limit and safe from overflow when added.</summary>
    private const long ExpansionCap = long.MaxValue / 4;
    private const int LinearDuplicateLimit = 8;

    private readonly YamlDocument _document;
    private readonly YamlLoadOptions _options;
    private YamlNode[] _nodes = new YamlNode[64];
    private int _nodeCount;
    private Frame[] _frames = new Frame[16];
    private int _frameCount;
    private Dictionary<string, YamlNode>? _anchors;
    private YamlDocumentPart? _part;
    private int _documentNodes;
    private long _expansion;
    private string?[] _identities = new string?[16];
    private Dictionary<string, int>? _seen;

    private struct Frame {
        public YamlNode Node;
        public int First;
        public long Expanded;
    }

    public YamlTreeBuilder(YamlDocument document, YamlLoadOptions options) {
        _document = document;
        _options = options;
    }

    public void Run(YamlReader reader) {
        while (reader.Read()) {
            ref readonly YamlEvent ev = ref reader.CurrentEvent;
            switch (ev.Type) {
                case YamlTokenType.DocumentStart:
                    _part = new YamlDocumentPart(_document, reader.DocumentSchema, reader.DocumentVersion, reader.CopyTagDirectives(), !ev.Implicit, ev.Start, ev.End);
                    _anchors?.Clear();
                    _expansion = 0;
                    _documentNodes = 0;
                    break;
                case YamlTokenType.DocumentEnd:
                    _part!._root = _nodes[0];
                    _nodes[0] = null!;
                    _nodeCount = 0;
                    _part._end = ev.End;
                    _part.HasEndMarker = !ev.Implicit;
                    _document.AddPart(_part);
                    break;
                case YamlTokenType.Scalar: {
                    var scalar = new YamlScalar(_part!, in ev);
                    if (ev.IsKey) scalar.SetDecodedValue(reader.GetString());
                    SetProperties(scalar, reader, in ev, open: false);
                    Count(in ev);
                    Push(scalar, 1);
                    break;
                }
                case YamlTokenType.Alias:
                    AddAlias(reader, in ev);
                    break;
                case YamlTokenType.MappingStart:
                    Open(new YamlMapping(_part!, in ev), reader, in ev);
                    break;
                case YamlTokenType.SequenceStart:
                    Open(new YamlSequence(_part!, in ev), reader, in ev);
                    break;
                case YamlTokenType.MappingEnd:
                case YamlTokenType.SequenceEnd:
                    Close(in ev);
                    break;
            }
        }
    }

    private void Open(YamlNode node, YamlReader reader, in YamlEvent ev) {
        SetProperties(node, reader, in ev, open: true);
        Count(in ev);
        node._parent = _frameCount == 0 ? null : _frames[_frameCount - 1].Node;
        if (_frameCount == _frames.Length) Array.Resize(ref _frames, _frames.Length * 2);
        _frames[_frameCount++] = new Frame { Node = node, First = _nodeCount, Expanded = 1 };
    }

    private void Close(in YamlEvent ev) {
        Frame frame = _frames[--_frameCount];
        _frames[_frameCount].Node = null!;
        YamlNode node = frame.Node;
        int count = _nodeCount - frame.First;
        if (node is YamlMapping mapping) {
            var entries = new YamlMappingEntry[count / 2];
            for (int i = 0; i < entries.Length; i++) entries[i] = new YamlMappingEntry(_nodes[frame.First + 2 * i], _nodes[frame.First + 2 * i + 1]);
            mapping._entries = entries;
            mapping._count = entries.Length;
            CheckDuplicates(mapping);
        } else {
            var sequence = (YamlSequence)node;
            var items = new YamlNode[count];
            Array.Copy(_nodes, frame.First, items, 0, count);
            sequence._items = items;
            sequence._count = count;
        }
        Array.Clear(_nodes, frame.First, count);
        _nodeCount = frame.First;
        node._end = ev.End;
        if (node._properties is { } properties) {
            properties.ExpandedSize = frame.Expanded;
            node._state &= ~YamlNodeState.Open;
        }
        PushChild(node, frame.Expanded);
    }

    private void AddAlias(YamlReader reader, in YamlEvent ev) {
        string name = reader.Anchor!;
        YamlPosition position = Position(in ev);
        if (_anchors is null || !_anchors.TryGetValue(name, out YamlNode? target)) {
            throw new YamlException(YamlErrorCodes.UndefinedAlias, "The alias *" + name + " has no anchor before it in its document.", position);
        }
        int limit = _options.MaxAliasExpansion;
        // An alias inside the collection it refers to makes the resolved graph infinite.
        if ((target._state & YamlNodeState.Open) != 0) throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxAliasExpansion), limit, position);
        long size = target._properties!.ExpandedSize;
        _expansion = Saturate(_expansion + size);
        if (_expansion > limit) throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxAliasExpansion), limit, position);
        Count(in ev);
        Push(new YamlAlias(_part!, ev.Start, ev.End, name, target), size);
    }

    private void SetProperties(YamlNode node, YamlReader reader, in YamlEvent ev, bool open) {
        if (ev.AnchorStart < 0 && !ev.HasTag) return;
        var properties = new YamlNodeProperties { Tag = reader.Tag, ExpandedSize = 1 };
        node._properties = properties;
        if (ev.AnchorStart < 0) return;
        string anchor = reader.Anchor!;
        properties.Anchor = anchor;
        if (open) node._state |= YamlNodeState.Open;
        (_anchors ??= new Dictionary<string, YamlNode>(StringComparer.Ordinal))[anchor] = node;
    }

    private void Count(in YamlEvent ev) {
        if (++_documentNodes > _options.MaxNodes) {
            throw new YamlLimitExceededException(nameof(YamlLoadOptions.MaxNodes), _options.MaxNodes, Position(in ev));
        }
    }

    /// <summary>Adds a leaf to the innermost open collection, or makes it the document root.</summary>
    private void Push(YamlNode node, long expanded) {
        node._parent = _frameCount == 0 ? null : _frames[_frameCount - 1].Node;
        PushChild(node, expanded);
    }

    private void PushChild(YamlNode node, long expanded) {
        if (_nodeCount == _nodes.Length) Array.Resize(ref _nodes, _nodes.Length * 2);
        _nodes[_nodeCount++] = node;
        if (_frameCount > 0) {
            ref Frame parent = ref _frames[_frameCount - 1];
            parent.Expanded = Saturate(parent.Expanded + expanded);
        }
    }

    private static long Saturate(long value) => value > ExpansionCap ? ExpansionCap : value;

    private static YamlPosition Position(in YamlEvent ev) => new YamlPosition(ev.Start, ev.StartLine + 1, ev.StartColumn + 1);

    /// <summary>
    /// Finds keys that are equal as YAML values. Each key gets one identity string (its decoded text for strings, so the
    /// common case allocates nothing); equal identities are confirmed by <see cref="YamlKeyEquality.Equal"/>. Small
    /// mappings compare identities pairwise, larger ones through a reused dictionary.
    /// </summary>
    private void CheckDuplicates(YamlMapping mapping) {
        int count = mapping._count;
        if (count < 2) return;
        YamlMappingEntry[] entries = mapping._entries;
        if (_identities.Length < count) _identities = new string?[Math.Max(count, _identities.Length * 2)];
        string?[] identities = _identities;
        for (int i = 0; i < count; i++) identities[i] = YamlKeyEquality.Identity(entries[i].Key);
        if (count <= LinearDuplicateLimit) {
            for (int i = 1; i < count; i++) {
                for (int j = 0; j < i; j++) {
                    if (string.Equals(identities[i], identities[j], StringComparison.Ordinal) && YamlKeyEquality.Equal(entries[j].Key, entries[i].Key)) {
                        ReportDuplicate(entries[i].Key, entries[j].Key);
                        break;
                    }
                }
            }
        } else {
            Dictionary<string, int> seen = _seen ??= new Dictionary<string, int>(StringComparer.Ordinal);
            seen.Clear();
            for (int i = 0; i < count; i++) {
                string? identity = identities[i];
                if (identity is null || seen.TryGetValue(identity, out _)) {
                    // Collection keys and identity collisions are rare: confirm against every earlier key.
                    for (int j = 0; j < i; j++) {
                        if (string.Equals(identities[j], identity, StringComparison.Ordinal) && YamlKeyEquality.Equal(entries[j].Key, entries[i].Key)) {
                            ReportDuplicate(entries[i].Key, entries[j].Key);
                            break;
                        }
                    }
                }
                if (identity is not null && !seen.ContainsKey(identity)) seen.Add(identity, i);
            }
        }
        Array.Clear(identities, 0, count);
    }

    private void ReportDuplicate(YamlNode key, YamlNode first) {
        string text = YamlMapping.KeyText(key) ?? key.ToString();
        if (text.Length > 64) text = text.Substring(0, 61) + "...";
        YamlPosition position = _document.GetPosition(key._start);
        string message = "The mapping key '" + text + "' at " + position + " repeats the key at " + _document.GetPosition(first._start) + ".";
        if (_options.DuplicateKeys == YamlDuplicateKeyPolicy.Error) throw new YamlException(YamlErrorCodes.DuplicateKey, message, position);
        _document.AddDiagnostic(new OfficeDiagnostic(YamlErrorCodes.DuplicateKey, OfficeDiagnosticSeverity.Warning, message, position.ToString(), "YAML"));
    }
}
