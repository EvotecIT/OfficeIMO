using System.Collections;

namespace OfficeIMO.Yaml;

/// <summary>A sequence node: block (<c>- item</c>) or flow (<c>[a, b]</c>), in source order.</summary>
public sealed class YamlSequence : YamlNode, IReadOnlyList<YamlNode> {
    internal YamlNode[] _items = Array.Empty<YamlNode>();
    internal int _count;
    internal int _contentStart;

    internal YamlSequence(YamlDocumentPart part, in YamlEvent ev) : base(part, ev.Style, ev.Start, ev.End) {
        _contentStart = ev.End;
    }

    /// <summary>Gets the number of items.</summary>
    public int Count => _count;

    /// <summary>Gets the item at a zero-based index. Aliases are returned as <see cref="YamlAlias"/> nodes.</summary>
    /// <exception cref="ArgumentOutOfRangeException"><paramref name="index"/> is outside the sequence.</exception>
    public YamlNode this[int index] {
        get {
            if ((uint)index >= (uint)_count) throw new ArgumentOutOfRangeException(nameof(index), index, "The index is outside the sequence.");
            return _items[index];
        }
    }

    /// <summary>Returns an enumerator over the items that does not allocate.</summary>
    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<YamlNode> IEnumerable<YamlNode>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    internal int IndexOf(YamlNode item) {
        for (int i = 0; i < _count; i++) {
            if (ReferenceEquals(_items[i], item)) return i;
        }
        return -1;
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
        public YamlNode Current => _sequence._items[_index];

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public bool MoveNext() => ++_index < _sequence._count;

        /// <inheritdoc />
        public void Reset() => _index = -1;

        /// <inheritdoc />
        public void Dispose() { }
    }
}
