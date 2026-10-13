namespace OfficeIMO.OpenDocument;

/// <summary>Projects repeated ODF XML elements as a lazy logical read-only list.</summary>
internal sealed class OdfRepeatedElementCollection<T> : IReadOnlyList<T> {
    private const int MaximumLogicalItems = 1000000;
    private readonly IReadOnlyList<XElement> _elements;
    private readonly XName _repeatAttribute;
    private readonly Func<XElement, long, T> _factory;
    private readonly int _count;

    internal OdfRepeatedElementCollection(IReadOnlyList<XElement> elements, XName repeatAttribute,
        Func<XElement, long, T> factory) {
        _elements = elements;
        _repeatAttribute = repeatAttribute;
        _factory = factory;
        long count = 0;
        foreach (XElement element in elements) {
            count = AddCount(count, OdsRepeatModel.Read(element, repeatAttribute));
        }
        _count = (int)count;
    }

    /// <summary>Admits a repeat before adding it, keeping oversized counts out of overflow paths.</summary>
    internal static long AddCount(long count, long repeat) {
        if (repeat > MaximumLogicalItems - count)
            throw OdfTextCodec.Limit(nameof(MaximumLogicalItems), MaximumLogicalItems);
        return count + repeat;
    }

    public int Count => _count;

    public T this[int index] {
        get {
            if (index < 0 || index >= _count) throw new ArgumentOutOfRangeException(nameof(index));
            long current = 0;
            foreach (XElement element in _elements) {
                long repeat = OdsRepeatModel.Read(element, _repeatAttribute);
                if (index < current + repeat) return _factory(element, index - current);
                current += repeat;
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }
    }

    public IEnumerator<T> GetEnumerator() {
        foreach (XElement element in _elements) {
            long repeat = OdsRepeatModel.Read(element, _repeatAttribute);
            for (long offset = 0; offset < repeat; offset++) yield return _factory(element, offset);
        }
    }

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
}
