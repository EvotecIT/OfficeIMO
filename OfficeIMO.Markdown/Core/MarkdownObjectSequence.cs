namespace OfficeIMO.Markdown;

/// <summary>
/// Append-only object list stored in growing chunks, so recording a large tree never copies it.
/// </summary>
internal sealed class MarkdownObjectSequence {
    // Objects are stored in a struct so writing one into a chunk needs no array type check:
    // MarkdownObject is not sealed, so a MarkdownObject[] store checks the element's type every time.
    internal struct Slot {
        internal MarkdownObject Value;
    }

    private const int FirstChunkSize = 16;
    private const int MaxChunkSize = 4096;

    private readonly List<Slot[]> _chunks = new();
    private Slot[] _current = Array.Empty<Slot>();
    private int _currentCount;
    private int _completedCount;

    internal int Count => _completedCount + _currentCount;

    internal void Add(MarkdownObject item) {
        if (_currentCount == _current.Length) {
            _completedCount += _currentCount;
            _current = new Slot[_current.Length == 0 ? FirstChunkSize : Math.Min(_current.Length * 2, MaxChunkSize)];
            _currentCount = 0;
            _chunks.Add(_current);
        }

        _current[_currentCount++].Value = item;
    }

    internal void AddRange(Slot[] items, int count) {
        for (int i = 0; i < count; i++) {
            Add(items[i].Value);
        }
    }

    internal MarkdownObject[] ToArray() {
        var items = new MarkdownObject[Count];
        int index = 0;
        foreach (MarkdownObject item in this) {
            items[index++] = item;
        }

        return items;
    }

    /// <summary>Visits the items in the order they were added.</summary>
    public Enumerator GetEnumerator() => new(this);

    // Items must not be added while an enumeration is running.
    public struct Enumerator {
        private readonly MarkdownObjectSequence _sequence;
        private int _chunk;
        private Slot[] _items;
        private int _length;
        private int _index;

        internal Enumerator(MarkdownObjectSequence sequence) {
            _sequence = sequence;
            _chunk = -1;
            _items = Array.Empty<Slot>();
            _length = 0;
            _index = -1;
        }

        public MarkdownObject Current => _items[_index].Value;

        public bool MoveNext() {
            if (++_index < _length) {
                return true;
            }

            var chunks = _sequence._chunks;
            while (++_chunk < chunks.Count) {
                _items = chunks[_chunk];
                _length = ReferenceEquals(_items, _sequence._current) ? _sequence._currentCount : _items.Length;
                _index = 0;
                if (_length > 0) {
                    return true;
                }
            }

            return false;
        }
    }
}
