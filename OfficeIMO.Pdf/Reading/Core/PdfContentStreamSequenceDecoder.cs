using System.Text;
using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

/// <summary>Decodes the ordered streams that form one logical PDF content stream.</summary>
internal static class PdfContentStreamSequenceDecoder {
    internal static bool TryDecode(
        IReadOnlyList<PdfStream> streams,
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadLimits limits,
        bool enforcePageContentLimit,
        out string content) {
        Guard.NotNull(streams, nameof(streams));
        Guard.NotNull(objects, nameof(objects));
        Guard.NotNull(limits, nameof(limits));

        content = string.Empty;
        var builder = new StringBuilder();
        long decodedBytes = 0L;
        for (int index = 0; index < streams.Count; index++) {
            PdfStream stream = streams[index];
            if (!StreamDecoder.TryDecode(
                    stream.Dictionary,
                    stream.Data,
                    limits.MaxDecodedStreamBytes,
                    out byte[] decoded,
                    objects)) return false;

            decodedBytes = checked(decodedBytes + decoded.Length);
            if (enforcePageContentLimit && decodedBytes > limits.MaxPageContentBytes) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.PageContentBytes,
                    limits.MaxPageContentBytes,
                    decodedBytes);
            }
            builder.Append(PdfEncoding.Latin1GetString(decoded));
        }

        content = builder.ToString();
        return true;
    }
}
