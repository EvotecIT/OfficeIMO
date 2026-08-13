namespace OfficeIMO.Pdf;

internal sealed class PdfDecodedStreamBudget {
    private readonly int _maximumPerStream;
    private readonly long _maximumTotal;
    private readonly Dictionary<PdfStream, byte[]> _decoded = new();
    private long _used;

    internal PdfDecodedStreamBudget(PdfReadLimits limits) {
        _maximumPerStream = limits.MaxDecodedStreamBytes;
        _maximumTotal = limits.MaxTotalDecodedStreamBytes;
    }

    internal long UsedBytes => _used;

    internal byte[] Decode(PdfStream stream, Dictionary<int, PdfIndirectObject> objects) {
        if (_decoded.TryGetValue(stream, out byte[]? cached)) return cached;
        long remaining = _maximumTotal - _used;
        if (remaining <= 0) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _used + 1);
        }
        int maximumOutput = (int)Math.Min(_maximumPerStream, Math.Min(remaining, int.MaxValue));
        byte[] decoded;
        try {
            decoded = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects, maximumOutput);
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
