namespace OfficeIMO.Toml;

/// <summary>
/// A table: the document root, a <c>[table]</c>, an inline table, a table made by dotted keys, or one table of an array
/// of tables. Keys are unique and keep source order; new keys are appended.
/// </summary>
/// <remarks>
/// A table is one logical object even when TOML spreads it over the file: <c>[a]</c>, <c>[a.b]</c> and
/// <c>a.c = 1</c> lines all contribute to the same tree. Editing keeps that layout: new key/value pairs are written after
/// the table's last pair, and new sub-tables after the table's last section.
/// </remarks>
public sealed class TomlTable : TomlNode, IReadOnlyCollection<KeyValuePair<string, TomlNode>> {
    private const int LinearLookupLimit = 8;

    internal TomlTableKind _tableKind;
    internal string[] _keys = Array.Empty<string>();
    internal TomlNode[] _nodes = Array.Empty<TomlNode>();
    internal int _count;
    private Dictionary<string, int>? _index;
    /// <summary>Loaded entries removed by editing, kept so saving can delete their text.</summary>
    internal List<TomlNode>? _removed;
    /// <summary>The index of the table's header in <see cref="TomlDocument._headers"/>, or -1.</summary>
    internal int _section = -1;

    internal TomlTable(TomlDocument document, TomlTableKind kind) : base(document) {
        _tableKind = kind;
    }

    /// <inheritdoc />
    public override TomlNodeKind Kind => TomlNodeKind.Table;

    /// <summary>Gets how the table is defined in the source.</summary>
    public TomlTableKind TableKind => _tableKind;

    /// <summary>Gets the number of keys.</summary>
    public int Count => _count;

    /// <summary>Gets the keys in order.</summary>
    public IEnumerable<string> Keys {
        get {
            for (int i = 0; i < _count; i++) yield return _keys[i];
        }
    }

    /// <summary>Gets the node stored under a key.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="KeyNotFoundException">The table has no such key.</exception>
    public TomlNode this[string key] {
        get {
            if (TryGetValue(key, out TomlNode? node)) return node!;
            throw new KeyNotFoundException("The table has no key '" + key + "'.");
        }
    }

    /// <summary>Gets the node stored under a key.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    public bool TryGetValue(string key, out TomlNode? node) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        int index = IndexOf(key);
        node = index >= 0 ? _nodes[index] : null;
        return index >= 0;
    }

    /// <summary>Returns whether the table has a key.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    public bool ContainsKey(string key) => TryGetValue(key, out _);

    /// <summary>
    /// Adds a key with a value and returns the new node: a <see cref="TomlValue"/>, or a <see cref="TomlArray"/> or
    /// inline <see cref="TomlTable"/> for sequences and dictionaries. New strings follow the quoting of their siblings.
    /// </summary>
    /// <param name="key">The key. Any string is allowed; keys that are not bare are written quoted.</param>
    /// <param name="value">A string, number, boolean or date-time (see <see cref="TomlValue.SetValue"/>), a sequence for
    /// an array, or a dictionary with string keys for an inline table.</param>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> or <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException">The key exists, or <paramref name="value"/> is not representable in TOML.</exception>
    public TomlNode Add(string key, object value) {
        RequireNewKey(key);
        TomlNode node = TomlValueConverter.CreateNode(_document, value, nameof(value), TomlEditing.PreferredStringStyle(this));
        Append(key, node);
        return node;
    }

    /// <summary>
    /// Sets a key's value and returns the node now stored there. An existing value is updated in place with
    /// <see cref="TomlValue.SetValue"/>, keeping its spelling when equal and its style when possible; a missing key is
    /// added as with <see cref="Add"/>; an existing table or array is replaced.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> or <paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="value"/> is not representable in TOML.</exception>
    public TomlNode Set(string key, object value) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        int index = IndexOf(key);
        if (index < 0) return Add(key, value);
        TomlNode current = _nodes[index];
        if (current is TomlValue scalar && TomlValueConverter.IsScalar(value)) return scalar.SetValue(value);
        TomlNode node = TomlValueConverter.CreateNode(_document, value, nameof(value), TomlEditing.PreferredStringStyle(this));
        if (current.IsKeyValue) {
            // Written at the old value's position on the same line.
            TomlEditing.Adopt(current, node);
            node._parent = this;
            node._key = key;
            _nodes[index] = node;
            TomlEditing.MarkChanged(this);
            return node;
        }
        RemoveAt(index);
        Append(key, node);
        return node;
    }

    /// <summary>
    /// Adds an empty table and returns it. Under a document table it is written as a <c>[header]</c> section after this
    /// table's last section; inside an inline table it is an inline table.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="ArgumentException">The key exists.</exception>
    public TomlTable AddTable(string key) {
        RequireNewKey(key);
        var table = new TomlTable(_document, IsInlineScope() ? TomlTableKind.Inline : TomlTableKind.Header);
        Append(key, table);
        return table;
    }

    /// <summary>Adds an empty inline table (<c>key = { }</c>) and returns it.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="ArgumentException">The key exists.</exception>
    public TomlTable AddInlineTable(string key) {
        RequireNewKey(key);
        var table = new TomlTable(_document, TomlTableKind.Inline);
        Append(key, table);
        return table;
    }

    /// <summary>Adds an empty array (<c>key = []</c>) and returns it.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="ArgumentException">The key exists.</exception>
    public TomlArray AddArray(string key) {
        RequireNewKey(key);
        var array = new TomlArray(_document);
        Append(key, array);
        return array;
    }

    /// <summary>
    /// Adds an empty array of tables and returns it; add its tables with <see cref="TomlTableArray.AddTable"/>. Each table
    /// is written as a <c>[[key]]</c> section. An array of tables with no tables has no text and is not written.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    /// <exception cref="ArgumentException">The key exists.</exception>
    /// <exception cref="InvalidOperationException">This table is inside an inline table or array, where <c>[[ ]]</c> headers cannot appear.</exception>
    public TomlTableArray AddTableArray(string key) {
        RequireNewKey(key);
        if (IsInlineScope()) throw new InvalidOperationException("An array of tables cannot be added inside an inline table; add an array of inline tables instead.");
        var array = new TomlTableArray(_document);
        Append(key, array);
        return array;
    }

    /// <summary>
    /// Removes a key and returns whether it existed. On save its text is deleted: the key/value line with the comment
    /// lines directly above it, or a table's header sections with all their contents.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="key"/> is null.</exception>
    public bool Remove(string key) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        int index = IndexOf(key);
        if (index < 0) return false;
        RemoveAt(index);
        return true;
    }

    /// <summary>Returns an enumerator over the keys and nodes in order.</summary>
    public IEnumerator<KeyValuePair<string, TomlNode>> GetEnumerator() {
        for (int i = 0; i < _count; i++) yield return new KeyValuePair<string, TomlNode>(_keys[i], _nodes[i]);
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Returns whether new tables under this one must be inline: it is inline or inside an inline table or array.</summary>
    internal bool IsInlineScope() => _tableKind == TomlTableKind.Inline || InInlineContext();

    internal int IndexOf(string key) {
        if (_count <= LinearLookupLimit) {
            for (int i = 0; i < _count; i++) {
                if (string.Equals(_keys[i], key, StringComparison.Ordinal)) return i;
            }
            return -1;
        }
        Dictionary<string, int> index = _index ??= BuildIndex();
        return index.TryGetValue(key, out int found) ? found : -1;
    }

    private Dictionary<string, int> BuildIndex() {
        var index = new Dictionary<string, int>(_count, StringComparer.Ordinal);
        for (int i = 0; i < _count; i++) index[_keys[i]] = i;
        return index;
    }

    private void RequireNewKey(string key) {
        if (key is null) throw new ArgumentNullException(nameof(key));
        if (IndexOf(key) >= 0) throw new ArgumentException("The table already has the key '" + key + "'.", nameof(key));
    }

    private void Append(string key, TomlNode node) {
        AddEntry(key, node);
        TomlEditing.MarkChanged(this);
    }

    private void RemoveAt(int index) {
        TomlNode node = _nodes[index];
        Array.Copy(_keys, index + 1, _keys, index, _count - index - 1);
        Array.Copy(_nodes, index + 1, _nodes, index, _count - index - 1);
        _count--;
        _keys[_count] = null!;
        _nodes[_count] = null!;
        _index = null;
        TomlEditing.Detach(node, ref _removed);
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Adds an entry without change tracking; the builder and editing both use it.</summary>
    internal void AddEntry(string key, TomlNode node) {
        if (_count == _keys.Length) {
            int size = Math.Max(4, _count * 2);
            Array.Resize(ref _keys, size);
            Array.Resize(ref _nodes, size);
        }
        _keys[_count] = key;
        _nodes[_count] = node;
        _index?.Add(key, _count);
        _count++;
        node._parent = this;
        node._key = key;
    }
}
