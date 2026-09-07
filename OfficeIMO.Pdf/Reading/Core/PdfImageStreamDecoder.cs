using System.Threading;
namespace OfficeIMO.Pdf;

internal static class PdfImageStreamDecoder {
    internal static bool TryDecode(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] decoded,
        int maxDecodedBytes = PdfReadLimits.DefaultMaxDecodedStreamBytes, CancellationToken cancellationToken = default) {
        try {
            decoded = Filters.StreamDecoder.DecodeRequired(
                stream.Dictionary,
                stream.Data,
                objects,
                maxDecodedBytes, cancellationToken);
            return true;
        } catch (InvalidDataException) {
            decoded = Array.Empty<byte>();
            return false;
        }
    }
}
