namespace OfficeIMO.Toml;

/// <summary>
/// An array value in brackets: <c>[1, 2, 3]</c>, possibly spanning lines with comments. Items can be values, arrays and
/// inline tables, of mixed kinds.
/// </summary>
public sealed class TomlArray : TomlNode, IReadOnlyList<TomlNode> {
    internal TomlNode[] _items = Array.Empty<TomlNode>();
    internal int _count;
    /// <summary>Loaded items removed by editing, kept so saving can delete their text.</summary>
    internal List<TomlNode>? _removed => _document.RemovedFrom(this);

    internal TomlArray(TomlDocument document) : base(document) {
    }

    /// <inheritdoc />
    public override TomlNodeKind Kind => TomlNodeKind.Array;

    /// <summary>Gets the number of items.</summary>
    public int Count => _count;

    /// <summary>Gets the item at a zero-based position.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the array.</exception>
    public TomlNode this[int index] {
        get {
            if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the array.");
            return _items[index];
        }
    }

    /// <summary>
    /// Appends an item and returns it. On save the item is written after the last item in the array's own layout: on a
    /// new line with the same indentation in a multi-line array, otherwise after a comma on the same line.
    /// </summary>
    /// <param name="value">A string, number, boolean or date-time (see <see cref="TomlValue.SetValue"/>), a sequence for a
    /// nested array, or a dictionary with string keys for an inline table.</param>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="value"/> is not representable in TOML.</exception>
    public TomlNode Add(object value) => Insert(_count, value);

    /// <summary>Inserts an item at a zero-based position and returns it.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside 0..<see cref="Count"/>.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="value"/> is not representable in TOML.</exception>
    public TomlNode Insert(int index, object value) {
        if ((uint)index > (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the array.");
        TomlNode node = TomlValueConverter.CreateNode(_document, value, nameof(value), TomlEditing.PreferredStringStyle(this));
        InsertNode(index, node);
        return node;
    }

    /// <summary>Appends an empty inline table and returns it, for arrays of inline tables such as <c>authors</c>.</summary>
    public TomlTable AddInlineTable() {
        var table = new TomlTable(_document, TomlTableKind.Inline);
        InsertNode(_count, table);
        return table;
    }

    /// <summary>Appends an empty nested array and returns it.</summary>
    public TomlArray AddArray() {
        var array = new TomlArray(_document);
        InsertNode(_count, array);
        return array;
    }

    /// <summary>
    /// Replaces the item at a zero-based position and returns the item now there; an index equal to <see cref="Count"/>
    /// appends as <see cref="Add"/> does. A value assigned over a value is set in place with
    /// <see cref="TomlValue.SetValue"/>, keeping its style; other replacements write the new item at the old item's position.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside 0..<see cref="Count"/>.</exception>
    /// <exception cref="ArgumentNullException"><paramref name="value"/> is null.</exception>
    /// <exception cref="ArgumentException"><paramref name="value"/> is not representable in TOML.</exception>
    public TomlNode Set(int index, object value) {
        if (index == _count) return Insert(index, value);
        TomlNode current = this[index];
        if (current is TomlValue scalar && TomlValueConverter.IsScalar(value)) return scalar.SetValue(value);
        TomlNode node = TomlValueConverter.CreateNode(_document, value, nameof(value), TomlEditing.PreferredStringStyle(this));
        PutAt(index, node);
        return node;
    }

    /// <summary>Replaces the item at a position with a new node written at the old item's position.</summary>
    internal void PutAt(int index, TomlNode node) {
        TomlEditing.Adopt(_items[index], node);
        node._parent = this;
        _items[index] = node;
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Removes the item at a zero-based position. On save its text and one separating comma are deleted.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the array.</exception>
    public void RemoveAt(int index) {
        TomlNode node = this[index];
        Array.Copy(_items, index + 1, _items, index, _count - index - 1);
        _items[--_count] = null!;
        TomlEditing.Detach(this, node);
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Returns an enumerator over the items.</summary>
    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<TomlNode> IEnumerable<TomlNode>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Enumerates the items of an array in order without allocating.</summary>
    public struct Enumerator : IEnumerator<TomlNode> {
        private readonly TomlArray _array;
        private int _index;

        internal Enumerator(TomlArray array) {
            _array = array;
            _index = -1;
        }

        /// <inheritdoc />
        public TomlNode Current => _array._items[_index];

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _array._count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() {
        }
    }

    internal void InsertNode(int index, TomlNode node) {
        if (_count == _items.Length) Array.Resize(ref _items, Math.Max(4, _count * 2));
        if (index < _count) Array.Copy(_items, index, _items, index + 1, _count - index);
        _items[index] = node;
        _count++;
        node._parent = this;
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Appends a loaded item while building the model.</summary>
    internal void AddLoaded(TomlNode node) {
        if (_count == _items.Length) Array.Resize(ref _items, Math.Max(4, _count * 2));
        _items[_count++] = node;
        node._parent = this;
    }
}
