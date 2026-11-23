namespace OfficeIMO.Pdf;

/// <summary>Finds raw security markers during the existing sequential scan of a seekable input.</summary>
internal sealed class PdfSeekableRawMarkerIndex {
    // Recognition here is byte-for-byte. The security parser still interprets each match.
    internal const int Overlap = 8; // "startxref" is the longest marker.
    private const int MaximumCandidatesPerMarker = 4096;
    private readonly Marker[] _markers = {
        new("/Encrypt"), new("startxref"), new("/Prev"), new("/Root"), new("/Info")
    };
    private readonly int _length;
    private int _scannedLength;

    internal PdfSeekableRawMarkerIndex(int length) => _length = length;

    internal void ObserveBlock(int offset, int blockLength, ReadOnlySpan<byte> bytesWithOverlap) {
        if (offset != _scannedLength) return;
        foreach (Marker marker in _markers) {
            if (marker.Offsets is not { } offsets) continue;
            for (int position = 0; position < blockLength;) {
                int relative = bytesWithOverlap.Slice(position).IndexOf(marker.Bytes);
                if (relative < 0 || relative >= blockLength - position) break;
                position += relative;
                if (offsets.Count == MaximumCandidatesPerMarker) {
                    // Preserve exact behavior for dense inputs without retaining a large index.
                    marker.Offsets = null;
                    break;
                }
                offsets.Add(offset + position);
                position++;
            }
        }
        _scannedLength += blockLength;
    }

    internal bool TryFind(ReadOnlySpan<byte> needle, int start, out int index) {
        index = -1;
        if (_scannedLength != _length) return false;
        foreach (Marker marker in _markers) {
            if (!needle.SequenceEqual(marker.Bytes) || marker.Offsets is not { } offsets) continue;
            int candidate = offsets.BinarySearch(start);
            if (candidate < 0) candidate = ~candidate;
            if (candidate < offsets.Count) index = offsets[candidate];
            return true;
        }
        return false;
    }

    private sealed class Marker {
        internal Marker(string value) => Bytes = Encoding.ASCII.GetBytes(value);

        internal byte[] Bytes { get; }
        internal List<int>? Offsets { get; set; } = new();
    }
}
