namespace OfficeIMO.Pdf;

internal sealed class PdfDecodedStreamBudget {
    private readonly int _maximumPerStream;
    private readonly long _maximumTotal;
    private readonly Dictionary<PdfStream, byte[]> _decoded = new();
    private long _used;

    internal PdfDecodedStreamBudget(PdfReadLimits limits, long initialUsedBytes = 0) {
        _maximumPerStream = limits.MaxDecodedStreamBytes;
        _maximumTotal = limits.MaxTotalDecodedStreamBytes;
        if (initialUsedBytes < 0 || initialUsedBytes > _maximumTotal) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, initialUsedBytes);
        }
        _used = initialUsedBytes;
    }

    internal long UsedBytes => _used;

    internal byte[] Decode(PdfStream stream, Dictionary<int, PdfIndirectObject> objects) {
        return Decode(stream, objects, _maximumPerStream);
    }

    internal byte[] Decode(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, int maximumRequestedBytes) {
        return DecodeCore(stream, objects, maximumRequestedBytes, requireSupportedFilters: false);
    }

    internal byte[] DecodeRequired(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, int maximumRequestedBytes) {
        return DecodeCore(stream, objects, maximumRequestedBytes, requireSupportedFilters: true);
    }

    private byte[] DecodeCore(PdfStream stream, Dictionary<int, PdfIndirectObject> objects, int maximumRequestedBytes, bool requireSupportedFilters) {
        if (_decoded.TryGetValue(stream, out byte[]? cached)) {
            if (cached.Length > maximumRequestedBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumRequestedBytes, cached.Length);
            }
            return cached;
        }
        long remaining = _maximumTotal - _used;
        if (remaining <= 0) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _used + 1);
        }
        int maximumOutput = (int)Math.Min(_maximumPerStream, Math.Min(maximumRequestedBytes, Math.Min(remaining, int.MaxValue)));
        byte[] decoded;
        try {
            decoded = requireSupportedFilters
                ? Filters.StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumOutput)
                : Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, maximumOutput);
        } catch (PdfReadLimitException exception) when (
            exception.Kind == PdfReadLimitKind.DecodedStreamBytes && remaining < _maximumPerStream) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _maximumTotal + 1);
        }
        _used = checked(_used + decoded.LongLength);
        if (_used > _maximumTotal) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _used);
        }
        _decoded.Add(stream, decoded);
        return decoded;
    }
}
