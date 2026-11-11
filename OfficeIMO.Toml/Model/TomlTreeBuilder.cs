namespace OfficeIMO.Toml;

/// <summary>
/// Builds the model from reader tokens and applies the rules the reader cannot check alone: unique keys, and how tables
/// may be defined and extended by headers, dotted keys and inline tables. It is iterative, so deep input cannot
/// overflow the stack.
/// </summary>
internal sealed class TomlTreeBuilder {
    private readonly TomlDocument _document;
    private readonly TomlLoadOptions _options;
    private readonly KeyCache _cache = new KeyCache();
    private readonly List<Frame> _frames = new List<Frame>();
    private readonly Frame _top = new Frame();
    private int _nodes;

    public TomlTreeBuilder(TomlDocument document, TomlLoadOptions options) {
        _document = document;
        _options = options;
    }

    /// <summary>The key/value pair being read at one nesting level: its keys and the container they belong to.</summary>
    private sealed class Frame {
        public TomlNode Container = null!;
        public readonly List<string> Keys = new List<string>(4);
        public readonly List<int> KeyOffsets = new List<int>(4);
        public int KeyEnd;
    }

    public void Run(TomlReader reader) {
        TomlTable section = _document._root;
        _top.Container = section;
        while (reader.Read()) {
            TomlTokenType type = reader.TokenType;
            Frame frame = _frames.Count == 0 ? _top : _frames[_frames.Count - 1];
            switch (type) {
                case TomlTokenType.TableHeaderStart:
                case TomlTokenType.ArrayOfTablesHeaderStart:
                    section = ReadHeader(reader, type == TomlTokenType.ArrayOfTablesHeaderStart);
                    _top.Container = section;
                    break;
                case TomlTokenType.Key:
                    frame.Keys.Add(_cache.Get(reader));
                    frame.KeyOffsets.Add(reader.TokenStartOffset);
                    frame.KeyEnd = reader.TokenEndOffset;
                    break;
                case TomlTokenType.EndArray:
                case TomlTokenType.EndInlineTable:
                    TomlNode closed = frame.Container;
                    closed._end = reader.TokenEndOffset;
                    if (closed is TomlTable inline) inline._state |= TomlNodeState.Sealed;
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
        TomlNode node = CreateNode(reader, type);
        if (frame.Container is TomlArray array) {
            SetDepth(node, array, reader.TokenStartOffset);
            array.AddLoaded(node);
        } else {
            Place((TomlTable)frame.Container, frame, node);
        }
        if (node is TomlArray || node is TomlTable) {
            Frame inner = _frames.Count < _pool.Count ? _pool[_frames.Count] : NewFrame();
            inner.Container = node;
            inner.Keys.Clear();
            inner.KeyOffsets.Clear();
            _frames.Add(inner);
        }
    }

    private readonly List<Frame> _pool = new List<Frame>();

    private Frame NewFrame() {
        var frame = new Frame();
        _pool.Add(frame);
        return frame;
    }

    private TomlNode CreateNode(TomlReader reader, TomlTokenType type) {
        if (++_nodes > _options.MaxNodes) throw new TomlLimitExceededException(nameof(TomlLoadOptions.MaxNodes), _options.MaxNodes, reader.Start);
        TomlNode node;
        switch (type) {
            case TomlTokenType.StartArray:
                node = new TomlArray(_document);
                break;
            case TomlTokenType.StartInlineTable:
                node = new TomlTable(_document, TomlTableKind.Inline);
                break;
            default:
                var value = new TomlValue(_document, KindOf(type));
                if (type == TomlTokenType.String) {
                    value._style = reader.StringStyle;
                } else if (type == TomlTokenType.Integer) {
                    ReadOnlySpan<char> raw = reader.ValueSpan;
                    if (raw.Length > 2 && raw[0] == '0') value._base = raw[1] == 'x' ? TomlIntegerBase.Hexadecimal : raw[1] == 'o' ? TomlIntegerBase.Octal : raw[1] == 'b' ? TomlIntegerBase.Binary : TomlIntegerBase.Decimal;
                }
                node = value;
                break;
        }
        node._start = reader.TokenStartOffset;
        node._end = reader.TokenEndOffset;
        return node;
    }

    private static TomlNodeKind KindOf(TomlTokenType type) => type switch {
        TomlTokenType.String => TomlNodeKind.String,
        TomlTokenType.Integer => TomlNodeKind.Integer,
        TomlTokenType.Float => TomlNodeKind.Float,
        TomlTokenType.Boolean => TomlNodeKind.Boolean,
        _ => TomlNodeKind.DateTime
    };

    /// <summary>Places a key/value pair in its table, creating the dotted tables of a dotted key.</summary>
    private void Place(TomlTable owner, Frame frame, TomlNode node) {
        List<string> keys = frame.Keys;
        TomlTable table = owner;
        for (int i = 0; i < keys.Count - 1; i++) {
            int index = table.IndexOf(keys[i]);
            if (index >= 0) {
                if (table._nodes[index] is TomlTable dotted && dotted._tableKind == TomlTableKind.Dotted) {
                    table = dotted;
                    continue;
                }
                throw Error(TomlErrorCodes.InvalidTable, "The dotted key goes through '" + keys[i] + "', which is already defined and cannot be extended here.", frame.KeyOffsets[i]);
            }
            var created = new TomlTable(_document, TomlTableKind.Dotted) { _state = TomlNodeState.Loaded };
            CountNode(frame.KeyOffsets[i]);
            SetDepth(created, table, frame.KeyOffsets[i]);
            table.AddEntry(keys[i], created);
            table = created;
        }
        string last = keys[keys.Count - 1];
        if (table.IndexOf(last) >= 0) throw Error(TomlErrorCodes.DuplicateKey, "The key '" + last + "' is already defined.", frame.KeyOffsets[keys.Count - 1]);
        SetDepth(node, table, frame.KeyOffsets[keys.Count - 1]);
        table.AddEntry(last, node);
        node._keyStart = frame.KeyOffsets[0];
        node._keyEnd = frame.KeyEnd;
        keys.Clear();
        frame.KeyOffsets.Clear();
    }

    /// <summary>Reads a header's keys and returns the table it opens.</summary>
    private TomlTable ReadHeader(TomlReader reader, bool arrayOfTables) {
        int start = reader.TokenStartOffset;
        Frame frame = _top;
        frame.Keys.Clear();
        frame.KeyOffsets.Clear();
        while (reader.Read() && reader.TokenType == TomlTokenType.Key) {
            frame.Keys.Add(_cache.Get(reader));
            frame.KeyOffsets.Add(reader.TokenStartOffset);
        }
        int end = reader.TokenEndOffset;
        List<string> keys = frame.Keys;
        TomlTable table = _document._root;
        for (int i = 0; i < keys.Count - 1; i++) {
            int index = table.IndexOf(keys[i]);
            if (index < 0) {
                var created = new TomlTable(_document, TomlTableKind.Implicit) { _state = TomlNodeState.Loaded };
                CountNode(frame.KeyOffsets[i]);
                SetDepth(created, table, frame.KeyOffsets[i]);
                table.AddEntry(keys[i], created);
                table = created;
                continue;
            }
            switch (table._nodes[index]) {
                case TomlTable existing when existing._tableKind != TomlTableKind.Inline:
                    table = existing;
                    break;
                case TomlTableArray tables:
                    table = tables._items[tables._items.Count - 1];
                    break;
                default:
                    throw Error(TomlErrorCodes.InvalidTable, "The header goes through '" + keys[i] + "', which is not a table that can be extended.", frame.KeyOffsets[i]);
            }
        }
        string last = keys[keys.Count - 1];
        int lastOffset = frame.KeyOffsets[keys.Count - 1];
        int found = table.IndexOf(last);
        TomlTable result;
        if (arrayOfTables) {
            TomlTableArray tables;
            if (found < 0) {
                tables = new TomlTableArray(_document) { _state = TomlNodeState.Loaded };
                CountNode(lastOffset);
                SetDepth(tables, table, lastOffset);
                table.AddEntry(last, tables);
            } else if (table._nodes[found] is TomlTableArray existing) {
                tables = existing;
            } else {
                throw Error(TomlErrorCodes.InvalidArrayOfTables, "'" + last + "' is already defined and is not an array of tables.", lastOffset);
            }
            result = new TomlTable(_document, TomlTableKind.ArrayItem);
            CountNode(lastOffset);
            SetDepth(result, tables, lastOffset);
            tables.AddLoaded(result);
        } else if (found < 0) {
            result = new TomlTable(_document, TomlTableKind.Header);
            CountNode(lastOffset);
            SetDepth(result, table, lastOffset);
            table.AddEntry(last, result);
        } else if (table._nodes[found] is TomlTable existing && existing._tableKind == TomlTableKind.Implicit) {
            existing._tableKind = TomlTableKind.Header;
            result = existing;
        } else {
            throw Error(TomlErrorCodes.InvalidTable, "The table '" + last + "' is already defined.", lastOffset);
        }
        result._start = start;
        result._end = end;
        result._section = _document._headers.Count;
        _document._headers.Add(result);
        keys.Clear();
        frame.KeyOffsets.Clear();
        return result;
    }

    private void CountNode(int offset) {
        if (++_nodes > _options.MaxNodes) throw new TomlLimitExceededException(nameof(TomlLoadOptions.MaxNodes), _options.MaxNodes, _document.GetPosition(offset));
    }

    private void SetDepth(TomlNode node, TomlNode parent, int offset) {
        int depth = parent._depth + 1;
        if (depth > _options.MaxDepth) throw new TomlLimitExceededException(nameof(TomlLoadOptions.MaxDepth), _options.MaxDepth, _document.GetPosition(offset));
        node._depth = depth;
    }

    private TomlException Error(string code, string message, int offset) => new TomlException(code, message, _document.GetPosition(offset));

    /// <summary>Reuses the strings of repeated keys (<c>name</c>, <c>version</c>, <c>dependencies</c>) instead of allocating each time.</summary>
    private sealed class KeyCache {
        private const int Size = 256;
        private const int MaxLength = 32;
        private readonly string?[] _entries = new string?[Size];

        public string Get(TomlReader reader) {
            ReadOnlySpan<char> raw = reader.ValueSpan;
            if (reader.StringStyle != TomlStringStyle.Bare || raw.Length > MaxLength) return reader.GetString();
            int hash = raw.Length;
            for (int i = 0; i < raw.Length; i++) hash = unchecked(hash * 31 + raw[i]);
            int slot = hash & (Size - 1);
            string? cached = _entries[slot];
            if (cached != null && raw.SequenceEqual(cached.AsSpan())) return cached;
            string created = raw.ToString();
            _entries[slot] = created;
            return created;
        }
    }
}
