namespace OfficeIMO.Pdf;

internal sealed class PdfContentOrderKey : IComparable<PdfContentOrderKey>, IEquatable<PdfContentOrderKey> {
    // Every span of page content gets a one-segment key, so that common depth keeps its segment inline
    // instead of allocating an array per span. Deeper keys (form content) keep all segments in the array.
    private readonly int[]? _segments;
    private readonly int _single;
    private readonly int _length;

    private PdfContentOrderKey(int[] segments) {
        _segments = segments;
        _length = segments.Length;
    }

    private PdfContentOrderKey(int single) {
        _single = single;
        _length = 1;
    }

    internal static PdfContentOrderKey Root { get; } = new PdfContentOrderKey(Array.Empty<int>());

    internal int Depth => _length;
    internal int? RootOperatorOffset => _length == 0 ? null : Segment(0);

    private int Segment(int index) => _segments is null ? _single : _segments[index];

    internal PdfContentOrderKey Append(int segment) {
        if (_length == 0) return new PdfContentOrderKey(segment);
        var segments = new int[_length + 1];
        for (int index = 0; index < _length; index++) segments[index] = Segment(index);
        segments[_length] = segment;
        return new PdfContentOrderKey(segments);
    }

    public int CompareTo(PdfContentOrderKey? other) {
        if (other == null) return 1;
        int commonLength = Math.Min(_length, other._length);
        for (int i = 0; i < commonLength; i++) {
            int comparison = Segment(i).CompareTo(other.Segment(i));
            if (comparison != 0) return comparison;
        }
        return _length.CompareTo(other._length);
    }

    public bool Equals(PdfContentOrderKey? other) => other != null && CompareTo(other) == 0;

    public override bool Equals(object? obj) => obj is PdfContentOrderKey other && Equals(other);

    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            for (int i = 0; i < _length; i++) {
                hash = (hash * 31) + Segment(i);
            }
            return hash;
        }
    }
}
