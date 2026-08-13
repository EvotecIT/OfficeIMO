namespace OfficeIMO.Pdf;

internal sealed class PdfDecodedStreamBudget {
    private readonly int _maximumPerStream;
    private readonly long _maximumTotal;
    private readonly Dictionary<PdfStream, DecodedEntry> _decoded = new();
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
        if (_decoded.TryGetValue(stream, out DecodedEntry? cached)) {
            if (cached.Bytes.Length > maximumRequestedBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maximumRequestedBytes, cached.Bytes.Length);
            }
            if (!requireSupportedFilters || cached.RequiredValidated) return cached.Bytes;
            long replacementAllowance = _maximumTotal - _used + cached.Bytes.LongLength;
            int replacementMaximum = (int)Math.Min(
                _maximumPerStream,
                Math.Min(maximumRequestedBytes, Math.Min(replacementAllowance, int.MaxValue)));
            byte[] required;
            try {
                required = Filters.StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, replacementMaximum);
            } catch (PdfReadLimitException exception) when (
                exception.Kind == PdfReadLimitKind.DecodedStreamBytes &&
                replacementAllowance < Math.Min(_maximumPerStream, (long)maximumRequestedBytes)) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _maximumTotal + 1);
            }
            long revisedUsed = checked(_used - cached.Bytes.LongLength + required.LongLength);
            if (revisedUsed > _maximumTotal) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, revisedUsed);
            }
            _used = revisedUsed;
            _decoded[stream] = new DecodedEntry(required, requiredValidated: true);
            return required;
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
            exception.Kind == PdfReadLimitKind.DecodedStreamBytes &&
            remaining < Math.Min(_maximumPerStream, (long)maximumRequestedBytes)) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _maximumTotal + 1);
        }
        _used = checked(_used + decoded.LongLength);
        if (_used > _maximumTotal) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _used);
        }
        _decoded.Add(stream, new DecodedEntry(decoded, requireSupportedFilters));
        return decoded;
    }

    private sealed class DecodedEntry {
        internal DecodedEntry(byte[] bytes, bool requiredValidated) {
            Bytes = bytes;
            RequiredValidated = requiredValidated;
        }

        internal byte[] Bytes { get; }
        internal bool RequiredValidated { get; }
    }
}
