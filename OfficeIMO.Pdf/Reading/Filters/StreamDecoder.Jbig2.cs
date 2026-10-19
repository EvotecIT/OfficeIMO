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
    /// Names the first rejected JBIG2 feature, decoding preceding filters through the bounded
    /// stream pipeline before scanning segment headers and JBIG2Globals without decoding regions.
    /// </summary>
    internal static string? FindUnsupportedJbig2Feature(PdfStream stream, Dictionary<int, PdfIndirectObject>? objects, int maximumBytes,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!stream.Dictionary.Items.TryGetValue("Filter", out PdfObject? declared) ||
            !TryGetFilterNames(declared, objects, out List<string> filters, cancellationToken)) return null;
        for (int index = 0; index < filters.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (GetFilterKind(filters[index]) != DecodeFilterKind.Jbig2) continue;
            byte[] data;
            if (index == 0) data = stream.GetData(cancellationToken);
            else if (!TryDecodeCore(stream, maximumBytes, out data, out PdfReadLimitException? limit, objects,
                         cancellationToken, stopBeforeFilter: index)) {
                if (limit != null) throw limit;
                return "malformed";
            }
            string? feature = InspectJbig2Headers(stream.Dictionary, index, data, objects, maximumBytes, cancellationToken);
            if (feature != null) return feature;
        }
        return null;
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
