namespace OfficeIMO.Toml;

/// <summary>
/// The read-only tree <see cref="TomlSerializer"/> reads text through. It records each node's kind, key and source span
/// in one pooled array of structs instead of building a <see cref="TomlDocument"/>, so reading a document into objects
/// allocates little beyond the objects themselves. The structural rules are the ones <see cref="TomlTreeBuilder"/>
/// applies. Input the rules reject, and any value that does not fit its type, stops the read with
/// <see cref="TomlReadFallbackException"/>; the serializer then reads the text again through <see cref="TomlDocument"/>,
/// which reports the error with its position and path.
/// </summary>
internal sealed class TomlDataTree {
    private const int LinearLookupLimit = 8;
    private const int RetainedItems = 1 << 16;

    [ThreadStatic] private static TomlDataTree? t_cached;

    internal string _source = string.Empty;
    internal Item[] _items = new Item[64];
    private int _count;
    private TomlLoadOptions _options = TomlLoadOptions.Default;
    private int _nodes;
    private Dictionary<(int Table, string Key), int>? _index;
    private readonly List<Frame> _frames = new List<Frame>();
    private readonly List<Frame> _pool = new List<Frame>();
    private readonly Frame _top = new Frame();

    /// <summary>One node. Containers link their children in order through <see cref="First"/> and <see cref="Next"/>.</summary>
    internal struct Item {
        public string? Key;
        public int Start;
        public int End;
        public int First;
        public int Last;
        public int Next;
        public int Count;
        public TomlNodeKind Kind;
        public TomlTableKind TableKind;
        public TomlStringStyle Style;
        public bool Indexed;
    }

    private sealed class Frame {
        public int Container;
        public int Depth;
        public readonly List<string> Keys = new List<string>(4);
    }

    /// <summary>The root table.</summary>
    public TomlReadNode Root => new TomlReadNode(this, 0);

    /// <summary>Takes this thread's tree, or a new one when it is in use, such as by a nested read.</summary>
    public static TomlDataTree Rent() {
        TomlDataTree? tree = t_cached;
        if (tree is null) return new TomlDataTree();
        t_cached = null;
        return tree;
    }

    /// <summary>Clears the tree and keeps it for the thread's next read, unless a very large input grew it.</summary>
    public static void Return(TomlDataTree tree) {
        Array.Clear(tree._items, 0, tree._count);
        tree._count = 0;
        tree._source = string.Empty;
        tree._index?.Clear();
        if (tree._items.Length <= RetainedItems) t_cached = tree;
    }

    /// <summary>
    /// Builds the tree. Returns false for input <see cref="TomlDocument.Parse"/> rejects, so the caller can let the
    /// document report it.
    /// </summary>
    public bool TryLoad(string text, TomlLoadOptions options) {
        if (options.MaxInputBytes is long max && text.Length > max) return false;
        _source = text;
        _options = options;
        _nodes = 0;
        _frames.Clear();
        Add(new Item { Kind = TomlNodeKind.Table, TableKind = TomlTableKind.Root, Start = -1, First = -1, Last = -1, Next = -1 });
        try {
            Run(new TomlReader(text.AsMemory(), options.ReaderOptions, CancellationToken.None));
            return true;
        } catch (TomlException) {
            return false;
        } catch (TomlReadFallbackException) {
            return false;
        }
    }

    private void Run(TomlReader reader) {
        _top.Container = 0;
        _top.Depth = 0;
        _top.Keys.Clear();
        while (reader.Read()) {
            TomlTokenType type = reader.TokenType;
            Frame frame = _frames.Count == 0 ? _top : _frames[_frames.Count - 1];
            switch (type) {
                case TomlTokenType.TableHeaderStart:
                case TomlTokenType.ArrayOfTablesHeaderStart:
                    _top.Container = ReadHeader(reader, type == TomlTokenType.ArrayOfTablesHeaderStart, out _top.Depth);
                    break;
                case TomlTokenType.Key:
                    frame.Keys.Add(reader.GetString());
                    break;
                case TomlTokenType.EndArray:
                case TomlTokenType.EndInlineTable:
                    _items[frame.Container].End = reader.TokenEndOffset;
                    _frames.RemoveAt(_frames.Count - 1);
                    break;
                case TomlTokenType.Comment:
                    break;
                default:
                    AddValue(reader, type, frame);
                    break;
            }
        }
    }

    private void AddValue(TomlReader reader, TomlTokenType type, Frame frame) {
        int node = CreateNode(reader, type);
        int depth;
        if (_items[frame.Container].Kind == TomlNodeKind.Array) {
            depth = CheckDepth(frame.Depth);
            Append(frame.Container, node);
        } else {
            depth = Place(frame.Container, frame, node);
        }
        TomlNodeKind kind = _items[node].Kind;
        if (kind == TomlNodeKind.Array || kind == TomlNodeKind.Table) {
            Frame inner = _frames.Count < _pool.Count ? _pool[_frames.Count] : NewFrame();
            inner.Container = node;
            inner.Depth = depth;
            inner.Keys.Clear();
            _frames.Add(inner);
        }
    }

    private Frame NewFrame() {
        var frame = new Frame();
        _pool.Add(frame);
        return frame;
    }

    private int CreateNode(TomlReader reader, TomlTokenType type) {
        CountNode();
        var item = new Item { Start = reader.TokenStartOffset, End = reader.TokenEndOffset, First = -1, Last = -1, Next = -1 };
        switch (type) {
            case TomlTokenType.StartArray:
                item.Kind = TomlNodeKind.Array;
                break;
            case TomlTokenType.StartInlineTable:
                item.Kind = TomlNodeKind.Table;
                item.TableKind = TomlTableKind.Inline;
                break;
            case TomlTokenType.String:
                item.Kind = TomlNodeKind.String;
                item.Style = reader.StringStyle;
                break;
            case TomlTokenType.Integer:
                item.Kind = TomlNodeKind.Integer;
                break;
            case TomlTokenType.Float:
                item.Kind = TomlNodeKind.Float;
                break;
            case TomlTokenType.Boolean:
                item.Kind = TomlNodeKind.Boolean;
                break;
            default:
                item.Kind = TomlNodeKind.DateTime;
                break;
        }
        return Add(item);
    }

    private int NewTable(TomlTableKind kind) {
        CountNode();
        return Add(new Item { Kind = TomlNodeKind.Table, TableKind = kind, Start = -1, First = -1, Last = -1, Next = -1 });
    }

    /// <summary>Places a key/value pair in its table, creating the dotted tables of a dotted key, and returns the node's depth.</summary>
    private int Place(int owner, Frame frame, int node) {
        List<string> keys = frame.Keys;
        int table = owner;
        int depth = frame.Depth;
        for (int i = 0; i < keys.Count - 1; i++) {
            int index = IndexOf(table, keys[i]);
            if (index >= 0) {
                if (_items[index].Kind == TomlNodeKind.Table && _items[index].TableKind == TomlTableKind.Dotted) {
                    table = index;
                    depth++;
                    continue;
                }
                throw new TomlReadFallbackException();
            }
            int created = NewTable(TomlTableKind.Dotted);
            depth = CheckDepth(depth);
            AddEntry(table, keys[i], created);
            table = created;
        }
        string last = keys[keys.Count - 1];
        if (IndexOf(table, last) >= 0) throw new TomlReadFallbackException();
        depth = CheckDepth(depth);
        AddEntry(table, last, node);
        keys.Clear();
        return depth;
    }

    /// <summary>Reads a header's keys and returns the table it opens, with that table's depth.</summary>
    private int ReadHeader(TomlReader reader, bool arrayOfTables, out int resultDepth) {
        List<string> keys = _top.Keys;
        keys.Clear();
        while (reader.Read() && reader.TokenType == TomlTokenType.Key) keys.Add(reader.GetString());
        int table = 0;
        int depth = 0;
        for (int i = 0; i < keys.Count - 1; i++) {
            int index = IndexOf(table, keys[i]);
            if (index < 0) {
                int created = NewTable(TomlTableKind.Implicit);
                depth = CheckDepth(depth);
                AddEntry(table, keys[i], created);
                table = created;
                continue;
            }
            ref Item existing = ref _items[index];
            if (existing.Kind == TomlNodeKind.Table && existing.TableKind != TomlTableKind.Inline) {
                table = index;
                depth++;
            } else if (existing.Kind == TomlNodeKind.TableArray) {
                table = existing.Last;
                depth += 2;
            } else {
                throw new TomlReadFallbackException();
            }
        }
        string last = keys[keys.Count - 1];
        int found = IndexOf(table, last);
        int result;
        if (arrayOfTables) {
            int tables;
            if (found < 0) {
                CountNode();
                tables = Add(new Item { Kind = TomlNodeKind.TableArray, Start = -1, First = -1, Last = -1, Next = -1 });
                CheckDepth(depth);
                AddEntry(table, last, tables);
            } else if (_items[found].Kind == TomlNodeKind.TableArray) {
                tables = found;
            } else {
                throw new TomlReadFallbackException();
            }
            result = NewTable(TomlTableKind.ArrayItem);
            resultDepth = CheckDepth(depth + 1);
            Append(tables, result);
        } else if (found < 0) {
            result = NewTable(TomlTableKind.Header);
            resultDepth = CheckDepth(depth);
            AddEntry(table, last, result);
        } else if (_items[found].Kind == TomlNodeKind.Table && _items[found].TableKind == TomlTableKind.Implicit) {
            _items[found].TableKind = TomlTableKind.Header;
            result = found;
            resultDepth = depth + 1;
        } else {
            throw new TomlReadFallbackException();
        }
        keys.Clear();
        return result;
    }

    private int Add(Item item) {
        if (_count == _items.Length) Array.Resize(ref _items, _items.Length * 2);
        _items[_count] = item;
        return _count++;
    }

    /// <summary>Links a child at the end of a container.</summary>
    private void Append(int container, int child) {
        ref Item parent = ref _items[container];
        if (parent.Last < 0) parent.First = child;
        else _items[parent.Last].Next = child;
        parent.Last = child;
        parent.Count++;
    }

    private void AddEntry(int table, string key, int child) {
        _items[child].Key = key;
        Append(table, child);
        ref Item parent = ref _items[table];
        if (parent.Indexed) {
            _index!.Add((table, key), child);
        } else if (parent.Count > LinearLookupLimit) {
            Dictionary<(int, string), int> index = _index ??= new Dictionary<(int, string), int>();
            for (int i = parent.First; i >= 0; i = _items[i].Next) index.Add((table, _items[i].Key!), i);
            parent.Indexed = true;
        }
    }

    internal int IndexOf(int table, string key) {
        if (_items[table].Indexed) return _index!.TryGetValue((table, key), out int found) ? found : -1;
        for (int i = _items[table].First; i >= 0; i = _items[i].Next) {
            if (string.Equals(_items[i].Key, key, StringComparison.Ordinal)) return i;
        }
        return -1;
    }

    private void CountNode() {
        if (++_nodes > _options.MaxNodes) throw new TomlReadFallbackException();
    }

    private int CheckDepth(int parentDepth) {
        int depth = parentDepth + 1;
        if (depth > _options.MaxDepth) throw new TomlReadFallbackException();
        return depth;
    }
}

/// <summary>
/// Stops a read through <see cref="TomlDataTree"/> so the serializer reads the text again through
/// <see cref="TomlDocument"/>, which reports what went wrong with its position and path.
/// </summary>
internal sealed class TomlReadFallbackException : Exception {
}
