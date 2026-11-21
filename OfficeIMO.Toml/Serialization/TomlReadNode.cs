namespace OfficeIMO.Toml;

/// <summary>
/// The node a converter reads: a <see cref="TomlNode"/> of a loaded or edited document, or a node of the
/// <see cref="TomlDataTree"/> the serializer reads text through.
/// </summary>
internal readonly struct TomlReadNode {
    private readonly TomlNode? _node;
    private readonly TomlDataTree? _tree;
    private readonly int _index;

    public TomlReadNode(TomlNode node) {
        _node = node;
        _tree = null;
        _index = 0;
    }

    internal TomlReadNode(TomlDataTree tree, int index) {
        _node = null;
        _tree = tree;
        _index = index;
    }

    public TomlNodeKind Kind => _node is not null ? _node.Kind : _tree!._items[_index].Kind;

    public bool IsTable => Kind == TomlNodeKind.Table;

    /// <summary>Whether the node is a string, number, boolean or date-time.</summary>
    public bool IsValue => Kind >= TomlNodeKind.String;

    /// <summary>The key of a table entry.</summary>
    public string? Key => _node is not null ? _node._key : _tree!._items[_index].Key;

    /// <summary>The number of entries or items of a container.</summary>
    public int Count => _node switch {
        null => _tree!._items[_index].Count,
        TomlTable table => table._count,
        TomlArray array => array._count,
        TomlTableArray tables => tables._items.Count,
        _ => 0
    };

    /// <summary>
    /// The document node, for reporting an error. A node of the data tree has none, so asking stops the read and the
    /// serializer reads the text again through <see cref="TomlDocument"/> to report it.
    /// </summary>
    public TomlNode Node => _node ?? throw new TomlReadFallbackException();

    /// <summary>The entries of a table, or the items of an array or array of tables, in order.</summary>
    public Children GetChildren() => _node is null ? new Children(_tree!, _tree!._items[_index].First) : new Children(_node);

    /// <summary>Finds a table entry by key.</summary>
    public bool TryGetChild(string key, out TomlReadNode child) {
        if (_node is null) {
            int index = _tree!.IndexOf(_index, key);
            child = index >= 0 ? new TomlReadNode(_tree, index) : default;
            return index >= 0;
        }
        if (((TomlTable)_node).TryGetValue(key, out TomlNode? found)) {
            child = new TomlReadNode(found!);
            return true;
        }
        child = default;
        return false;
    }

    /// <summary>The source text of an unedited value; edited values have only their decoded form.</summary>
    public bool TryGetRaw(out ReadOnlySpan<char> raw) {
        if (_node is null) {
            ref TomlDataTree.Item item = ref _tree!._items[_index];
            raw = _tree._source.AsSpan(item.Start, item.End - item.Start);
            return true;
        }
        var value = (TomlValue)_node;
        if (value._value is null && value._start >= 0) {
            raw = value._document._source.AsSpan(value._start, value._end - value._start);
            return true;
        }
        raw = default;
        return false;
    }

    public TomlStringStyle StringStyle => _node is not null ? ((TomlValue)_node)._style : _tree!._items[_index].Style;

    /// <summary>The decoded value of a document node that <see cref="TryGetRaw"/> has no text for.</summary>
    public object DecodedValue => ((TomlValue)_node!).Value;

    /// <summary>Enumerates a container's children without allocating.</summary>
    public struct Children {
        private readonly TomlDataTree? _tree;
        private readonly TomlNode? _container;
        private int _next;
        private TomlReadNode _current;

        internal Children(TomlDataTree tree, int first) {
            _tree = tree;
            _container = null;
            _next = first;
            _current = default;
        }

        internal Children(TomlNode container) {
            _tree = null;
            _container = container;
            _next = 0;
            _current = default;
        }

        public TomlReadNode Current => _current;

        public Children GetEnumerator() => this;

        public bool MoveNext() {
            if (_tree is not null) {
                if (_next < 0) return false;
                _current = new TomlReadNode(_tree, _next);
                _next = _tree._items[_next].Next;
                return true;
            }
            TomlNode? item = _container switch {
                TomlTable table => _next < table._count ? table._nodes[_next] : null,
                TomlArray array => _next < array._count ? array._items[_next] : null,
                TomlTableArray tables => _next < tables._items.Count ? tables._items[_next] : null,
                _ => null
            };
            if (item is null) return false;
            _next++;
            _current = new TomlReadNode(item);
            return true;
        }
    }
}
