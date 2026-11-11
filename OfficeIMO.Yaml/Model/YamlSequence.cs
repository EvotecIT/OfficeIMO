using System.Collections;

namespace OfficeIMO.Yaml;

/// <summary>A sequence node: block (<c>- item</c>) or flow (<c>[a, b]</c>), in source order.</summary>
public sealed class YamlSequence : YamlNode, IReadOnlyList<YamlNode> {
    /// <summary>The items; null for a loaded sequence until code first reaches them (see <see cref="Items"/>).</summary>
    internal YamlNode[]? _items = Array.Empty<YamlNode>();
    internal int _count;
    internal int _contentStart;
    /// <summary>The loaded items as (start, end) source pairs, kept from the first insertion or removal so saving can find their source text.</summary>
    internal int[]? _original;

    internal YamlSequence(YamlDocumentPart part, YamlStyle style, int start, int end, int contentStart, int count) : base(part, style, start, end) {
        _contentStart = contentStart;
        _count = count;
        _items = count == 0 ? Array.Empty<YamlNode>() : null;
    }

    /// <summary>The items, created from the document's node table on first use.</summary>
    internal YamlNode[] Items => Volatile.Read(ref _items) ?? _part!._document._table!.MaterializeItems(this);

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
            return Items[index];
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

    internal int IndexOf(YamlNode item) {
        int known = item._index;
        if ((uint)known < (uint)_count && ReferenceEquals(Items[known], item)) return known;
        for (int i = 0; i < _count; i++) {
            if (ReferenceEquals(Items[i], item)) return i;
        }
        return -1;
    }

    /// <summary>Records each item's position from <paramref name="from"/> on, so lookups take constant time.</summary>
    internal void Renumber(int from) {
        for (int i = from; i < _count; i++) Items[i]._index = i;
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
        public YamlNode Current => _sequence.Items[_index];

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _sequence._count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() { }
    }
}
