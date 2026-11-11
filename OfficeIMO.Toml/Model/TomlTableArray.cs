namespace OfficeIMO.Toml;

/// <summary>
/// An array of tables: the tables of the <c>[[name]]</c> headers with one name, in source order, such as the
/// <c>[[package]]</c> entries of a lock file.
/// </summary>
public sealed class TomlTableArray : TomlNode, IReadOnlyList<TomlTable> {
    internal readonly List<TomlTable> _items = new List<TomlTable>();
    /// <summary>Loaded tables removed by editing, kept so saving can delete their sections.</summary>
    internal List<TomlNode>? _removed;

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
        table._parent = this;
        _items.Add(table);
        TomlEditing.MarkChanged(this);
        return table;
    }

    /// <summary>Removes the table at a zero-based position. On save its sections are deleted.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the array.</exception>
    public void RemoveAt(int index) {
        TomlTable table = this[index];
        _items.RemoveAt(index);
        TomlEditing.Detach(table, ref _removed);
        TomlEditing.MarkChanged(this);
    }

    /// <summary>Returns an enumerator over the tables.</summary>
    public IEnumerator<TomlTable> GetEnumerator() => _items.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    internal void AddLoaded(TomlTable table) {
        table._parent = this;
        _items.Add(table);
    }
}
