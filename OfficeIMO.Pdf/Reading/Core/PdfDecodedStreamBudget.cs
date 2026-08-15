namespace OfficeIMO.Pdf;

internal sealed class PdfDecodedStreamBudget {
    private readonly int _maximumPerStream;
    private readonly long _maximumTotal;
    private readonly Dictionary<PdfStream, DecodedEntry> _decoded = new();
    private readonly Dictionary<PdfStream, RequiredDecodeFailure> _requiredFailures = new();
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
            long revalidationRemaining = _maximumTotal - _used;
            if (revalidationRemaining <= 0) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _used + 1);
            }
            int revalidationMaximum = (int)Math.Min(
                _maximumPerStream,
                Math.Min(maximumRequestedBytes, Math.Min(revalidationRemaining, int.MaxValue)));
            byte[] required;
            try {
                ThrowCachedRequiredFailure(stream, revalidationMaximum);
                required = DecodeRequiredAndCacheFailure(stream, objects, revalidationMaximum);
            } catch (PdfReadLimitException exception) when (
                exception.Kind == PdfReadLimitKind.DecodedStreamBytes &&
                revalidationRemaining < Math.Min(_maximumPerStream, (long)maximumRequestedBytes)) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.TotalDecodedStreamBytes, _maximumTotal, _maximumTotal + 1);
            }
            long revisedUsed = checked(_used + required.LongLength);
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
            if (requireSupportedFilters) ThrowCachedRequiredFailure(stream, maximumOutput);
            decoded = requireSupportedFilters
                ? DecodeRequiredAndCacheFailure(stream, objects, maximumOutput)
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
        bool requiredValidated = requireSupportedFilters ||
            Filters.StreamDecoder.HasNoEffectiveFilters(stream.Dictionary, objects);
        _decoded.Add(stream, new DecodedEntry(decoded, requiredValidated));
        return decoded;
    }

    private byte[] DecodeRequiredAndCacheFailure(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maximumOutput) {
        try {
            byte[] decoded = Filters.StreamDecoder.DecodeRequired(stream.Dictionary, stream.Data, objects, maximumOutput);
            _requiredFailures.Remove(stream);
            return decoded;
        } catch (PdfReadLimitException exception) when (exception.Kind == PdfReadLimitKind.DecodedStreamBytes) {
            _requiredFailures[stream] = RequiredDecodeFailure.ForLimit(maximumOutput, exception.Actual);
            throw;
        } catch (InvalidDataException exception) {
            _requiredFailures[stream] = RequiredDecodeFailure.ForInvalidData(exception.Message);
            throw;
        }
    }

    private void ThrowCachedRequiredFailure(PdfStream stream, int maximumOutput) {
        if (!_requiredFailures.TryGetValue(stream, out RequiredDecodeFailure? failure)) return;
        if (failure.Message != null) throw new InvalidDataException(failure.Message);
        if (maximumOutput <= failure.MaximumOutput) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.DecodedStreamBytes,
                maximumOutput,
                Math.Max(failure.Actual, (long)maximumOutput + 1L));
        }
    }

    private sealed class DecodedEntry {
        internal DecodedEntry(byte[] bytes, bool requiredValidated) {
            Bytes = bytes;
            RequiredValidated = requiredValidated;
        }

        internal byte[] Bytes { get; }
        internal bool RequiredValidated { get; }
    }

    private sealed class RequiredDecodeFailure {
        private RequiredDecodeFailure(int maximumOutput, long actual, string? message) {
            MaximumOutput = maximumOutput;
            Actual = actual;
            Message = message;
        }

        internal int MaximumOutput { get; }
        internal long Actual { get; }
        internal string? Message { get; }

        internal static RequiredDecodeFailure ForLimit(int maximumOutput, long actual) =>
            new(maximumOutput, actual, message: null);

        internal static RequiredDecodeFailure ForInvalidData(string message) =>
            new(0, 0, message);
    }
}
