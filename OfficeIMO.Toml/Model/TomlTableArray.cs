namespace OfficeIMO.Toml;

/// <summary>
/// An array of tables: the tables of the <c>[[name]]</c> headers with one name, in source order, such as the
/// <c>[[package]]</c> entries of a lock file.
/// </summary>
public sealed class TomlTableArray : TomlNode, IReadOnlyList<TomlTable> {
    internal readonly List<TomlTable> _items = new List<TomlTable>();
    /// <summary>Loaded tables removed by editing, kept so saving can delete their sections.</summary>
    internal List<TomlNode>? _removed => _document.RemovedFrom(this);

    internal TomlTableArray(TomlDocument document) : base(document) {
    }

    /// <inheritdoc />
    public override TomlNodeKind Kind => TomlNodeKind.TableArray;

    /// <summary>Gets the number of tables.</summary>
    public int Count => _items.Count;

    /// <summary>Gets the table at a zero-based position.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the array.</exception>
    public TomlTable this[int index] {
        get {
            if ((uint)index >= (uint)_items.Count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the array of tables.");
            return _items[index];
        }
    }

    /// <summary>
    /// Appends an empty table and returns it. On save it is written as a new <c>[[name]]</c> section after the last
    /// section of the array's last table.
    /// </summary>
    public TomlTable AddTable() {
        var table = new TomlTable(_document, TomlTableKind.ArrayItem);
        AddItem(table);
        return table;
    }

    /// <summary>Appends a new table of this array's kind, with its contents.</summary>
    internal void AddItem(TomlTable table) {
        table._parent = this;
        _items.Add(table);
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Removes the table at a zero-based position. On save its sections are deleted.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the array.</exception>
    public void RemoveAt(int index) {
        TomlTable table = this[index];
        _items.RemoveAt(index);
        TomlEditing.Detach(this, table);
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Returns an enumerator over the tables.</summary>
    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<TomlTable> IEnumerable<TomlTable>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Enumerates the tables in order without allocating.</summary>
    public struct Enumerator : IEnumerator<TomlTable> {
        private readonly List<TomlTable> _items;
        private int _index;

        internal Enumerator(TomlTableArray tables) {
            _items = tables._items;
            _index = -1;
        }

        /// <inheritdoc />
        public TomlTable Current => _items[_index];

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _items.Count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() {
        }
    }

    internal void AddLoaded(TomlTable table) {
        table._parent = this;
        _items.Add(table);
    }
}
