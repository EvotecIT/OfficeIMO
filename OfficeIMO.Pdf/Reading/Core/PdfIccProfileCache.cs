using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfIccProfileCache {
    private static readonly ConditionalWeakTable<PdfStream, CacheEntry> Entries = new ConditionalWeakTable<PdfStream, CacheEntry>();

    internal static bool TryRead(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        out OfficeIccColorProfile? profile) {
        CacheEntry entry = Entries.GetValue(stream, key => Decode(key, objects, maxDecodedBytes));
        if (entry.DecodedLength > maxDecodedBytes) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maxDecodedBytes, entry.DecodedLength);
        }
        profile = entry.Profile;
        return profile != null;
    }

    private static CacheEntry Decode(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        if (!PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] bytes, maxDecodedBytes)) {
            return new CacheEntry(0, null);
        }
        OfficeIccColorProfile.TryCreate(bytes, out OfficeIccColorProfile? profile);
        return new CacheEntry(bytes.LongLength, profile);
    }

    private sealed class CacheEntry {
        internal CacheEntry(long decodedLength, OfficeIccColorProfile? profile) {
            DecodedLength = decodedLength;
            Profile = profile;
        }
        internal long DecodedLength { get; }
        internal OfficeIccColorProfile? Profile { get; }
    }
}
