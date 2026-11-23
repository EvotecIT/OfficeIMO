namespace OfficeIMO.Pdf;

/// <summary>One bounded, cancellable read cache for a canonical parser pass.</summary>
internal sealed class PdfSeekableCursor {
    private const int BlockSize = 65536 - PdfSeekableRawMarkerIndex.Overlap;
    private readonly PdfSeekableInput _input;
    private readonly CancellationToken _cancellationToken;
    private readonly byte[] _buffer;
    private int _start = -1;
    private int _count;
    private PdfSeekableRawMarkerIndex? _rawMarkers;

    internal PdfSeekableCursor(PdfSeekableInput input, CancellationToken cancellationToken) {
        _input = input;
        _cancellationToken = cancellationToken;
        _buffer = new byte[Math.Min(BlockSize + PdfSeekableRawMarkerIndex.Overlap, input.Length)];
    }

    internal PdfSeekableInput Input => _input;

    internal void BeginRawSecurityScan() => _rawMarkers ??= new PdfSeekableRawMarkerIndex(_input.Length);

    internal bool TryFindRawSecurityMarker(ReadOnlySpan<byte> marker, int start, out int index) {
        index = -1;
        return _rawMarkers?.TryFind(marker, start, out index) == true;
    }

    internal byte ReadByte(int index) {
        if ((uint)(index - _start) >= (uint)_count) Fill(index);
        return _buffer[index - _start];
    }

    internal ReadOnlySpan<byte> ReadRange(int start, int length) {
        if (length == 0) return ReadOnlySpan<byte>.Empty;
        if ((uint)(start - _start) >= (uint)_count) Fill(start);
        if (length <= _count - (start - _start)) return _buffer.AsSpan(start - _start, length);
        return _input.ReadAt(start, length, _cancellationToken);
    }

    internal int IndexOf(byte value, int start, int count) {
        int end = start + count;
        while (start < end) {
            if ((uint)(start - _start) >= (uint)_count) Fill(start);
            int length = Math.Min(end - start, _count - (start - _start));
            int found = PdfByteSearch.IndexOf(_buffer, value, start - _start, length);
            if (found >= 0) return _start + found;
            start += length;
        }
        return -1;
    }

    internal int IndexOf(ReadOnlySpan<byte> needle, int start, int count) {
        if (needle.IsEmpty) return start;
        int end = start + count;
        while (start <= end - needle.Length) {
            if ((uint)(start - _start) >= (uint)_count) Fill(start);
            int blockEnd = Math.Min(end, _start + _count);
            int found = _buffer.AsSpan(start - _start, blockEnd - start).IndexOf(needle);
            if (found >= 0) return start + found;

            // The only remaining candidates begin in this block and end in the next.
            int boundaryStart = Math.Max(start, blockEnd - needle.Length + 1);
            int boundaryEnd = Math.Min(blockEnd, end - needle.Length + 1);
            for (int candidate = boundaryStart; candidate < boundaryEnd; candidate++) {
                int matched = 0;
                while (matched < needle.Length && ReadByte(candidate + matched) == needle[matched]) matched++;
                if (matched == needle.Length) return candidate;
            }
            start = blockEnd;
        }
        return -1;
    }

    internal int LastIndexOf(string needle, int start, int count, bool ignoreCase) {
        int end = start + count;
        if (needle.Length == 0) return end;
        if (!ignoreCase) {
            Span<byte> marker = needle.Length <= 64 ? stackalloc byte[needle.Length] : new byte[needle.Length];
            for (int index = 0; index < needle.Length; index++) marker[index] = (byte)needle[index];
            while (end - start >= marker.Length) {
                int last = end - 1;
                if ((uint)(last - _start) >= (uint)_count) Fill(last);
                int blockStart = Math.Max(start, _start);
                int found = _buffer.AsSpan(blockStart - _start, end - blockStart).LastIndexOf(marker);
                if (found >= 0) return blockStart + found;

                int boundaryStart = Math.Max(start, blockStart - marker.Length + 1);
                int boundaryEnd = Math.Min(blockStart - 1, end - marker.Length);
                for (int candidate = boundaryEnd; candidate >= boundaryStart; candidate--) {
                    int matched = 0;
                    while (matched < marker.Length && ReadByte(candidate + matched) == marker[matched]) matched++;
                    if (matched == marker.Length) return candidate;
                }
                end = blockStart;
            }
            return -1;
        }

        for (int candidate = end - needle.Length; candidate >= start; candidate--) {
            int matched = 0;
            while (matched < needle.Length && PdfByteContentText.SameCharacter(ReadByte(candidate + matched), needle[matched], ignoreCase)) {
                matched++;
            }
            if (matched == needle.Length) return candidate;
        }
        return -1;
    }
    private void Fill(int index) {
        if ((uint)index >= (uint)_input.Length) throw new ArgumentOutOfRangeException(nameof(index));
        int start = index - index % BlockSize;
        int count = Math.Min(_buffer.Length, _input.Length - start);
        _start = -1;
        _count = 0;
        _input.ReadAt(start, _buffer, 0, count, _cancellationToken);
        _start = start;
        _count = count;
        _rawMarkers?.ObserveBlock(start, Math.Min(BlockSize, count), _buffer.AsSpan(0, count));
    }
}
