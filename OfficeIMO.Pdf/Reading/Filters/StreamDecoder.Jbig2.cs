using System.Threading;

namespace OfficeIMO.Pdf.Filters;

internal static partial class StreamDecoder {
    [ThreadStatic] private static HashSet<PdfStream>? _activeJbig2Globals;

    private static byte[] DecodeJbig2Globals(PdfStream stream, Dictionary<int, PdfIndirectObject>? objects,
        int maximumBytes, CancellationToken cancellationToken) {
        HashSet<PdfStream> active = _activeJbig2Globals ??= new HashSet<PdfStream>();
        if (active.Count >= 32 || !active.Add(stream)) throw new InvalidDataException("Cyclic or excessively nested JBIG2Globals.");
        try { return Decode(stream, objects, maximumBytes, cancellationToken); }
        finally { active.Remove(stream); }
    }

    private static byte[] DecodeJbig2(PdfDictionary dictionary, int index, byte[] data,
        Dictionary<int, PdfIndirectObject>? objects, int maximumBytes, CancellationToken cancellationToken) {
        PdfDictionary parameters = GetDecodeParms(dictionary, index, objects, cancellationToken) ?? new PdfDictionary();
        byte[]? globals = null;
        if (parameters.Items.TryGetValue("JBIG2Globals", out PdfObject? globalsObject)) {
            if (ResolveObject(globalsObject, objects, cancellationToken) is not PdfStream globalsStream) {
                throw new InvalidDataException("JBIG2Globals must be a stream.");
            }
            globals = DecodeJbig2Globals(globalsStream, objects, maximumBytes, cancellationToken);
        }
        int width = ReadIntegerParameter(dictionary, "Width", 0, objects, cancellationToken);
        int height = ReadIntegerParameter(dictionary, "Height", 0, objects, cancellationToken);
        return Jbig2Decoder.Decode(data, globals, width, height, maximumBytes, cancellationToken);
    }

    /// <summary>
    /// Names the first rejected JBIG2 feature, inspecting each JBIG2 input while decoding
    /// the filter chain once through the bounded stream pipeline.
    /// </summary>
    internal static string? FindUnsupportedJbig2Feature(PdfStream stream, Dictionary<int, PdfIndirectObject>? objects, int maximumBytes,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? declared) ||
            !TryGetFilterNames(declared, objects, out List<string> filters, cancellationToken)) return null;
        int lastJbig2Index = -1;
        for (int index = 0; index < filters.Count; index++) {
            if ((index & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (GetFilterKind(filters[index]) == DecodeFilterKind.Jbig2) lastJbig2Index = index;
        }
        if (lastJbig2Index < 0) return null;
        const int maximumDiagnosticFilterStages = 256;
        if (filters.Count > maximumDiagnosticFilterStages) return "filter-chain-too-long";
        string? feature = null;
        if (!TryDecodeCore(stream, maximumBytes, out _, out PdfReadLimitException? limit, objects,
                cancellationToken, beforeFilter: (index, current, original, offset, length) => {
                    if (GetFilterKind(filters[index]) == DecodeFilterKind.Jbig2)
                        feature = InspectJbig2Headers(stream.Dictionary, index,
                            current ?? GetExactData(original, offset, length, cancellationToken),
                            objects, maximumBytes, cancellationToken);
                    return feature == null && index < lastJbig2Index;
                })) {
            if (limit != null) throw limit;
            return "malformed";
        }
        return feature;
    }

    private static string? InspectJbig2Headers(PdfDictionary dictionary, int filterIndex, byte[] data,
        Dictionary<int, PdfIndirectObject>? objects, int maximumBytes, CancellationToken cancellationToken) {
        PdfDictionary? parameters = GetDecodeParms(dictionary, filterIndex, objects, cancellationToken);
        if (parameters != null && parameters.Items.TryGetValue("JBIG2Globals", out PdfObject? globalsObject)) {
            if (ResolveObject(globalsObject, objects, cancellationToken) is not PdfStream globalsStream) return "malformed";
            byte[] globals;
            try {
                globals = DecodeJbig2Globals(globalsStream, objects, maximumBytes, cancellationToken);
            } catch (Exception exception) when (exception is NotSupportedException or InvalidDataException || exception is IOException && exception is not PdfReadLimitException) {
                return "malformed";
            }
            string? globalFeature = Jbig2Decoder.FindUnsupportedFeature(globals, cancellationToken);
            if (globalFeature != null) return globalFeature;
        }
        return Jbig2Decoder.FindUnsupportedFeature(data, cancellationToken);
    }
}
