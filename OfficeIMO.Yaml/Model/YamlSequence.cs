using System.Collections;

namespace OfficeIMO.Yaml;

/// <summary>A sequence node: block (<c>- item</c>) or flow (<c>[a, b]</c>), in source order.</summary>
public sealed class YamlSequence : YamlNode, IReadOnlyList<YamlNode> {
    /// <summary>
    /// The items; null for a loaded sequence until code first reaches them. While <see cref="_slots"/> is set, an item
    /// whose object was not created yet is null and its record is in the slot of the same position.
    /// </summary>
    internal YamlNode?[]? _items = Array.Empty<YamlNode>();
    /// <summary>The record of each item while some items have no object yet; null once all have.</summary>
    internal int[]? _slots;
    internal int _count;
    internal int _contentStart;
    /// <summary>The loaded items as (start, end) source pairs, kept from the first insertion or removal so saving can find their source text.</summary>
    internal int[]? _original;

    internal YamlSequence(YamlDocumentPart part, YamlStyle style, int start, int end, int contentStart, int count) : base(part, style, start, end) {
        _contentStart = contentStart;
        _count = count;
        _items = count == 0 ? Array.Empty<YamlNode>() : null;
    }

    private YamlNodeTable Table => _part!._document._table!;

    /// <summary>The items with their objects, creating the object of every item that has none yet.</summary>
    internal YamlNode[] Items {
        get {
            YamlNode?[] items = Volatile.Read(ref _items) ?? Table.CreateSlots(this);
            if (Volatile.Read(ref _slots) is not null) Table.LoadItems(this);
            return items!;
        }
    }

    /// <summary>The item array, whose items may still have no objects (see <see cref="IsLoaded"/>).</summary>
    internal YamlNode?[] Slots => Volatile.Read(ref _items) ?? Table.CreateSlots(this);

    /// <summary>The item at a position, creating the object of that item only.</summary>
    internal YamlNode ItemAt(int index) => Volatile.Read(ref (Volatile.Read(ref _items) ?? Table.CreateSlots(this))[index]) ?? Table.LoadItem(this, index);

    /// <summary>Whether the item at a position has its object; items without one are untouched source.</summary>
    internal bool IsLoaded(int index) => (Volatile.Read(ref _items) ?? Table.CreateSlots(this))[index] is not null;

    /// <summary>Where the item starts in the source; reads no object for an item that has none.</summary>
    internal int ItemStart(int index) => IsLoaded(index) ? _items![index]!._start : Table.Records[_slots![index]].Start;

    /// <summary>Where the item ends in the source; reads no object for an item that has none.</summary>
    internal int ItemEnd(int index) => IsLoaded(index) ? _items![index]!._end : Table.Records[_slots![index]].End;

    /// <summary>Creates an empty block sequence. Inside a flow collection it is written as flow.</summary>
    public YamlSequence() : base(YamlStyle.Block) {
    }

    /// <summary>Creates an empty sequence with a style: <see cref="YamlStyle.Block"/> or <see cref="YamlStyle.Flow"/> (<c>[a, b]</c>).</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="style"/> is not a collection style.</exception>
    public YamlSequence(YamlStyle style) : base(YamlEdit.CheckCollectionStyle(style)) {
    }

    /// <summary>Gets the number of items.</summary>
    public int Count => _count;

    /// <summary>
    /// Gets the item at a zero-based index, or replaces it. Aliases are returned as <see cref="YamlAlias"/> nodes. A
    /// scalar set over a scalar updates it in place, keeping its anchor, tag, comments and, for an equal value, its text.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the sequence.</exception>
    public YamlNode this[int index] {
        get {
            if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the sequence.");
            return ItemAt(index);
        }
        set {
            if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the sequence.");
            YamlEdit.ReplaceValue(this, index, value ?? new YamlScalar((string?)null));
        }
    }

    /// <summary>Adds an item at the end, in the layout and quoting of the other items.</summary>
    /// <returns>The item, now in the sequence.</returns>
    /// <exception cref="InvalidOperationException"><paramref name="item"/> already belongs to a collection.</exception>
    public YamlNode Add(YamlNode? item) => Insert(_count, item);

    /// <summary>Inserts an item before the item at <paramref name="index"/>, or at the end when it equals <see cref="Count"/>.</summary>
    /// <returns>The item, now in the sequence.</returns>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside 0..<see cref="Count"/>.</exception>
    public YamlNode Insert(int index, YamlNode? item) {
        if ((uint)index > (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the sequence.");
        return YamlEdit.InsertItem(this, index, item ?? new YamlScalar((string?)null));
    }

    /// <summary>Removes the item at a position together with the comment lines above it, its lines and its trailing comment.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the sequence.</exception>
    /// <exception cref="InvalidOperationException">Aliases elsewhere refer to an anchor inside the item.</exception>
    public void RemoveAt(int index) {
        if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the sequence.");
        YamlEdit.RemoveChild(this, index);
    }

    /// <summary>Returns an enumerator over the items that does not allocate.</summary>
    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<YamlNode> IEnumerable<YamlNode>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Finds the position of an item; items without objects cannot be it and are not created.</summary>
    internal int IndexOf(YamlNode item) {
        YamlNode?[] items = Volatile.Read(ref _items) ?? Table.CreateSlots(this);
        int known = item._index;
        if ((uint)known < (uint)_count && ReferenceEquals(items[known], item)) return known;
        for (int i = 0; i < _count; i++) {
            if (ReferenceEquals(items[i], item)) return i;
        }
        return -1;
    }

    /// <summary>
    /// Records each item's position from <paramref name="from"/> on, so lookups take constant time. Items without
    /// objects get theirs when the objects are created.
    /// </summary>
    internal void Renumber(int from) {
        YamlNode?[] items = _items!;
        for (int i = from; i < _count; i++) {
            if (items[i] is { } item) item._index = i;
        }
    }

    /// <summary>Enumerates the items of a <see cref="YamlSequence"/>.</summary>
    public struct Enumerator : IEnumerator<YamlNode> {
        private readonly YamlSequence _sequence;
        private int _index;

        internal Enumerator(YamlSequence sequence) {
            _sequence = sequence;
            _index = -1;
        }

        /// <inheritdoc />
        public YamlNode Current => _sequence.ItemAt(_index);

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _sequence._count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() { }
    }
}
