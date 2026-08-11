using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static class PdfIccProfileCache {
    private static readonly ConditionalWeakTable<PdfStream, CacheSlot> Entries = new ConditionalWeakTable<PdfStream, CacheSlot>();

    internal static bool TryRead(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        out OfficeIccColorProfile? profile) {
        CacheEntry entry = GetEntry(stream, objects, maxDecodedBytes);
        profile = entry.Profile;
        return profile != null;
    }

    internal static bool TryReadBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        out byte[] bytes) {
        CacheEntry entry = GetEntry(stream, objects, maxDecodedBytes);
        bytes = entry.Bytes;
        return entry.Decoded;
    }

    private static CacheEntry GetEntry(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        CacheSlot slot = Entries.GetValue(stream, _ => new CacheSlot());
        lock (slot.Sync) {
            CacheEntry? entry = slot.Entry;
            if (entry == null) {
                entry = Decode(stream, objects, maxDecodedBytes);
                slot.Entry = entry;
            }
            if (entry.DecodedLength > maxDecodedBytes) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.DecodedStreamBytes, maxDecodedBytes, entry.DecodedLength);
            }
            return entry;
        }
    }

    private static CacheEntry Decode(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        if (!PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] bytes, maxDecodedBytes)) {
            return new CacheEntry(Array.Empty<byte>(), decoded: false, null);
        }
        OfficeIccColorProfile.TryCreate(bytes, out OfficeIccColorProfile? profile);
        return new CacheEntry(bytes, decoded: true, profile);
    }

    private sealed class CacheEntry {
        internal CacheEntry(byte[] bytes, bool decoded, OfficeIccColorProfile? profile) {
            Bytes = bytes;
            Decoded = decoded;
            Profile = profile;
        }
        internal byte[] Bytes { get; }
        internal bool Decoded { get; }
        internal long DecodedLength => Bytes.LongLength;
        internal OfficeIccColorProfile? Profile { get; }
    }

    private sealed class CacheSlot {
        internal object Sync { get; } = new object();
        internal CacheEntry? Entry { get; set; }
    }
}
