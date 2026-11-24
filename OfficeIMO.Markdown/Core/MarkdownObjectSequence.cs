namespace OfficeIMO.Markdown;

/// <summary>
/// Append-only object list stored in growing chunks, so recording a large tree never copies it.
/// </summary>
internal sealed class MarkdownObjectSequence {
    private const int FirstChunkSize = 16;
    private const int MaxChunkSize = 4096;

    private readonly List<MarkdownObject[]> _chunks = new();
    private MarkdownObject[] _current = Array.Empty<MarkdownObject>();
    private int _currentCount;
    private int _completedCount;

    internal int Count => _completedCount + _currentCount;

    internal void Add(MarkdownObject item) {
        if (_currentCount == _current.Length) {
            _completedCount += _currentCount;
            _current = new MarkdownObject[_current.Length == 0 ? FirstChunkSize : Math.Min(_current.Length * 2, MaxChunkSize)];
            _currentCount = 0;
            _chunks.Add(_current);
        }

        _current[_currentCount++] = item;
    }

    internal void AddRange(List<MarkdownObject> items) {
        for (int i = 0; i < items.Count; i++) {
            Add(items[i]);
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

    public struct Enumerator {
        private readonly MarkdownObjectSequence _sequence;
        private int _chunk;
        private int _index;

        internal Enumerator(MarkdownObjectSequence sequence) {
            _sequence = sequence;
            _chunk = 0;
            _index = -1;
        }

        public MarkdownObject Current => _sequence._chunks[_chunk][_index];

        public bool MoveNext() {
            var chunks = _sequence._chunks;
            while (_chunk < chunks.Count) {
                int length = ReferenceEquals(chunks[_chunk], _sequence._current) ? _sequence._currentCount : chunks[_chunk].Length;
                if (++_index < length) {
                    return true;
                }

                _chunk++;
                _index = -1;
            }

            return false;
        }
    }
}
