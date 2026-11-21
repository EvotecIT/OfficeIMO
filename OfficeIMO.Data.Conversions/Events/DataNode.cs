namespace OfficeIMO.Data.Conversions;

/// <summary>
/// A buffered value, used only where a mapping needs lookahead: YAML anchors replayed by aliases, merge keys, XML
/// sibling grouping and table rows. Streaming routes never build a whole document from these.
/// </summary>
internal sealed class DataNode {
    private DataNode(DataTokenType type) {
        Type = type;
    }

    /// <summary><see cref="DataTokenType.StartObject"/>, <see cref="DataTokenType.StartArray"/> or <see cref="DataTokenType.Scalar"/>.</summary>
    public DataTokenType Type { get; }
    public DataScalarKind Kind { get; private set; }
    public string? Text { get; private set; }
    public List<KeyValuePair<string, DataNode>>? Properties { get; private set; }
    public List<DataNode>? Items { get; private set; }

    public static DataNode Object() => new DataNode(DataTokenType.StartObject) { Properties = new List<KeyValuePair<string, DataNode>>() };
    public static DataNode Array() => new DataNode(DataTokenType.StartArray) { Items = new List<DataNode>() };
    public static DataNode Scalar(DataScalarKind kind, string text) => new DataNode(DataTokenType.Scalar) { Kind = kind, Text = text };

    /// <summary>Counts this node and its descendants.</summary>
    public long Size {
        get {
            long size = 1;
            if (Properties != null) foreach (KeyValuePair<string, DataNode> property in Properties) size += property.Value.Size;
            if (Items != null) foreach (DataNode item in Items) size += item.Size;
            return size;
        }
    }

    public bool TryGet(string name, out DataNode? value) {
        if (Properties != null) {
            foreach (KeyValuePair<string, DataNode> property in Properties) {
                if (string.Equals(property.Key, name, StringComparison.Ordinal)) {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = null;
        return false;
    }

    /// <summary>Renders the node as compact flow YAML, for complex keys.</summary>
    public string ToFlowText() {
        var builder = new StringBuilder();
        AppendFlow(builder);
        return builder.ToString();
    }

    private void AppendFlow(StringBuilder builder) {
        if (Type == DataTokenType.Scalar) {
            builder.Append(Kind == DataScalarKind.Null ? "null" : Text);
            return;
        }

        if (Type == DataTokenType.StartArray) {
            builder.Append('[');
            for (int i = 0; i < Items!.Count; i++) {
                if (i > 0) builder.Append(", ");
                Items[i].AppendFlow(builder);
            }

            builder.Append(']');
            return;
        }

        builder.Append('{');
        for (int i = 0; i < Properties!.Count; i++) {
            if (i > 0) builder.Append(", ");
            builder.Append(Properties[i].Key).Append(": ");
            Properties[i].Value.AppendFlow(builder);
        }

        builder.Append('}');
    }
}

/// <summary>Builds a <see cref="DataNode"/> from a run of events.</summary>
internal sealed class DataNodeBuilder {
    private readonly Stack<DataNode> _open = new();
    private string? _pendingName;

    public DataNode? Result { get; private set; }

    public bool IsComplete => Result != null;

    public int Depth => _open.Count;

    /// <summary>Adds an already-built node as the next value.</summary>
    public void AddNode(DataNode node) {
        Attach(node);
        if (_open.Count == 0) Result = node;
    }

    public void Add(DataTokenType type, DataScalarKind kind, ReadOnlySpan<char> text) {
        switch (type) {
            case DataTokenType.StartObject: Open(DataNode.Object()); break;
            case DataTokenType.StartArray: Open(DataNode.Array()); break;
            case DataTokenType.EndObject:
            case DataTokenType.EndArray: Close(_open.Pop()); break;
            case DataTokenType.PropertyName: _pendingName = text.ToString(); break;
            case DataTokenType.Scalar: Close(DataNode.Scalar(kind, text.ToString())); break;
        }
    }

    private void Open(DataNode node) {
        Attach(node);
        _open.Push(node);
    }

    // Containers were attached when they opened; scalars attach as they close.
    private void Close(DataNode node) {
        if (node.Type == DataTokenType.Scalar) Attach(node);
        if (_open.Count == 0) Result = node;
    }

    private void Attach(DataNode node) {
        if (_open.Count == 0) return;
        DataNode parent = _open.Peek();
        if (parent.Type == DataTokenType.StartObject) {
            parent.Properties!.Add(new KeyValuePair<string, DataNode>(_pendingName ?? string.Empty, node));
            _pendingName = null;
        } else {
            parent.Items!.Add(node);
        }
    }
}

/// <summary>Replays a <see cref="DataNode"/> as events without recursion.</summary>
internal sealed class DataNodeReplay {
    private readonly Stack<Frame> _stack = new();
    private DataNode? _root;
    private bool _nameEmitted;

    private struct Frame {
        public DataNode Node;
        public int Index;
    }

    public DataNodeReplay(DataNode root) {
        _root = root;
    }

    /// <summary>Optional object members to replay instead of a whole node (used by merge keys).</summary>
    public static DataNodeReplay Members(IReadOnlyList<KeyValuePair<string, DataNode>> members) {
        DataNode holder = DataNode.Object();
        holder.Properties!.AddRange(members);
        var replay = new DataNodeReplay(holder);
        replay._root = null;
        replay._stack.Push(new Frame { Node = holder, Index = 0 });
        replay._membersOnly = true;
        return replay;
    }

    private bool _membersOnly;

    /// <summary>Produces the next event; false when the node is exhausted.</summary>
    public bool Next(out DataTokenType type, out DataScalarKind kind, out string? text) {
        kind = default;
        text = null;
        if (_root != null) {
            DataNode node = _root;
            _root = null;
            return Begin(node, out type, out kind, out text);
        }

        while (_stack.Count > 0) {
            Frame frame = _stack.Pop();
            DataNode parent = frame.Node;
            if (parent.Type == DataTokenType.StartObject) {
                if (frame.Index < parent.Properties!.Count) {
                    KeyValuePair<string, DataNode> member = parent.Properties[frame.Index];
                    if (!_nameEmitted) {
                        _nameEmitted = true;
                        _stack.Push(frame);
                        type = DataTokenType.PropertyName;
                        text = member.Key;
                        return true;
                    }

                    _nameEmitted = false;
                    frame.Index++;
                    _stack.Push(frame);
                    return Begin(member.Value, out type, out kind, out text);
                }

                if (_membersOnly && _stack.Count == 0) {
                    type = DataTokenType.None;
                    return false;
                }

                type = DataTokenType.EndObject;
                return true;
            }

            if (frame.Index < parent.Items!.Count) {
                DataNode item = parent.Items[frame.Index];
                frame.Index++;
                _stack.Push(frame);
                return Begin(item, out type, out kind, out text);
            }

            type = DataTokenType.EndArray;
            return true;
        }

        type = DataTokenType.None;
        return false;
    }

    private bool Begin(DataNode node, out DataTokenType type, out DataScalarKind kind, out string? text) {
        kind = default;
        text = null;
        if (node.Type == DataTokenType.Scalar) {
            type = DataTokenType.Scalar;
            kind = node.Kind;
            text = node.Text;
            return true;
        }

        _stack.Push(new Frame { Node = node, Index = 0 });
        type = node.Type;
        return true;
    }
}
