namespace OfficeIMO.Pdf;

internal sealed class PdfContentOrderKey : IComparable<PdfContentOrderKey>, IEquatable<PdfContentOrderKey> {
    private readonly int[] _segments;

    private PdfContentOrderKey(int[] segments) {
        _segments = segments;
    }

    internal static PdfContentOrderKey Root { get; } = new PdfContentOrderKey(Array.Empty<int>());

    internal int Depth => _segments.Length;

    internal PdfContentOrderKey Append(int segment) {
        var segments = new int[_segments.Length + 1];
        Array.Copy(_segments, segments, _segments.Length);
        segments[_segments.Length] = segment;
        return new PdfContentOrderKey(segments);
    }

    public int CompareTo(PdfContentOrderKey? other) {
        if (other == null) return 1;
        int commonLength = Math.Min(_segments.Length, other._segments.Length);
        for (int i = 0; i < commonLength; i++) {
            int comparison = _segments[i].CompareTo(other._segments[i]);
            if (comparison != 0) return comparison;
        }
        return _segments.Length.CompareTo(other._segments.Length);
    }

    public bool Equals(PdfContentOrderKey? other) => other != null && CompareTo(other) == 0;

    public override bool Equals(object? obj) => obj is PdfContentOrderKey other && Equals(other);

    public override int GetHashCode() {
        unchecked {
            int hash = 17;
            for (int i = 0; i < _segments.Length; i++) {
                hash = (hash * 31) + _segments[i];
            }
            return hash;
        }
    }
}
