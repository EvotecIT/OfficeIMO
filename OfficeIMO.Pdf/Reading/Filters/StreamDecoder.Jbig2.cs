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
    /// Names the first JBIG2 feature the managed decoder rejects when JBIG2Decode is the stream's first filter,
    /// scanning segment headers of the data and its JBIG2Globals without decoding regions.
    /// </summary>
    internal static string? FindUnsupportedJbig2Feature(PdfStream stream, Dictionary<int, PdfIndirectObject>? objects, int maximumBytes) {
        if (!TryGetDeclaredFilterNames(stream.Dictionary, objects, out IReadOnlyList<string> filters) || filters.Count == 0 ||
            GetFilterKind(filters[0]) != DecodeFilterKind.Jbig2) return null;
        PdfDictionary? parameters = GetDecodeParms(stream.Dictionary, 0, objects, CancellationToken.None);
        if (parameters != null && parameters.Items.TryGetValue("JBIG2Globals", out PdfObject? globalsObject)) {
            if (ResolveObject(globalsObject, objects, CancellationToken.None) is not PdfStream globalsStream) return "malformed";
            byte[] globals;
            try {
                globals = DecodeJbig2Globals(globalsStream, objects, maximumBytes, CancellationToken.None);
            } catch (Exception exception) when (exception is NotSupportedException || exception is IOException && exception is not PdfReadLimitException) {
                return "malformed";
            }
            string? globalFeature = Jbig2Decoder.FindUnsupportedFeature(globals);
            if (globalFeature != null) return globalFeature;
        }
        return Jbig2Decoder.FindUnsupportedFeature(stream.Data);
    }
}
