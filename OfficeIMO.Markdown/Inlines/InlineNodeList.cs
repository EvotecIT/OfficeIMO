namespace OfficeIMO.Markdown;

/// <summary>
/// Storage behind <see cref="InlineSequence.Nodes"/>. Most parsed sequences hold a single text run, so
/// the first node is kept in a field and only later nodes go to an array. The list is a live, read-only
/// view: changes made through <see cref="InlineSequence"/> show through references callers already hold.
/// </summary>
internal sealed class InlineNodeList : IList<IMarkdownInline>, IReadOnlyList<IMarkdownInline> {
    private IMarkdownInline? _first;
    private IMarkdownInline[]? _rest;
    private int _count;
    private readonly int _capacity;

    internal InlineNodeList(int capacity = 0) {
        _capacity = capacity;
    }

    public int Count => _count;

    public bool IsReadOnly => true;

    public IMarkdownInline this[int index] {
        get {
            if ((uint)index >= (uint)_count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            return index == 0 ? _first! : _rest![index - 1];
        }
        set => throw new NotSupportedException();
    }

    internal void Append(IMarkdownInline node) {
        if (_count == 0) {
            _first = node;
            _count = 1;
            return;
        }

        int restIndex = _count - 1;
        if (_rest == null) {
            _rest = new IMarkdownInline[Math.Max(_capacity - 1, 2)];
        } else if (restIndex == _rest.Length) {
            Array.Resize(ref _rest, _rest.Length * 2);
        }

        _rest[restIndex] = node;
        _count++;
    }

    internal void Reset() {
        _first = null;
        if (_rest != null && _count > 1) {
            Array.Clear(_rest, 0, _count - 1);
        }

        _count = 0;
    }

    public int IndexOf(IMarkdownInline item) {
        for (int i = 0; i < _count; i++) {
            if (EqualityComparer<IMarkdownInline>.Default.Equals(this[i], item)) {
                return i;
            }
        }

        return -1;
    }

    public bool Contains(IMarkdownInline item) => IndexOf(item) >= 0;

    public void CopyTo(IMarkdownInline[] array, int arrayIndex) {
        ArgumentNullException.ThrowIfNull(array);
        if (arrayIndex < 0 || array.Length - arrayIndex < _count) {
            throw new ArgumentOutOfRangeException(nameof(arrayIndex));
        }

        if (_count == 0) {
            return;
        }

        array[arrayIndex] = _first!;
        if (_count > 1) {
            Array.Copy(_rest!, 0, array, arrayIndex + 1, _count - 1);
        }
    }

    public Enumerator GetEnumerator() => new Enumerator(this);

    IEnumerator<IMarkdownInline> IEnumerable<IMarkdownInline>.GetEnumerator() => GetEnumerator();

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

    void IList<IMarkdownInline>.Insert(int index, IMarkdownInline item) => throw new NotSupportedException();

    void IList<IMarkdownInline>.RemoveAt(int index) => throw new NotSupportedException();

    void ICollection<IMarkdownInline>.Add(IMarkdownInline item) => throw new NotSupportedException();

    void ICollection<IMarkdownInline>.Clear() => throw new NotSupportedException();

    bool ICollection<IMarkdownInline>.Remove(IMarkdownInline item) => throw new NotSupportedException();

    internal struct Enumerator : IEnumerator<IMarkdownInline> {
        private readonly InlineNodeList _list;
        private int _index;

        internal Enumerator(InlineNodeList list) {
            _list = list;
            _index = -1;
        }

        public IMarkdownInline Current => _list[_index];

        object System.Collections.IEnumerator.Current => Current;

        public bool MoveNext() => ++_index < _list._count;

        public void Reset() => _index = -1;

        public void Dispose() {
        }
    }
}
