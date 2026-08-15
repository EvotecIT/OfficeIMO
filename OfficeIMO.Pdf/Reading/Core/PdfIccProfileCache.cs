using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Caller-owned aggregate retention budget for distinct parsed ICC profiles.</summary>
internal sealed class PdfIccProfileRetentionBudget {
    private readonly int _maximumRetainedBytes;
    private readonly Dictionary<(PdfStream Stream, PdfIccProfileCacheRepresentation Representation), long> _charges =
        new Dictionary<(PdfStream, PdfIccProfileCacheRepresentation), long>();
    private long _retainedBytes;

    internal PdfIccProfileRetentionBudget(int maximumRetainedBytes) {
        _maximumRetainedBytes = Math.Max(1, maximumRetainedBytes);
    }

    internal void Charge(
        PdfStream stream,
        PdfIccProfileCacheRepresentation representation,
        long retainedLength) {
        lock (_charges) {
            var key = (stream, representation);
            _charges.TryGetValue(key, out long priorCharge);
            if (retainedLength <= priorCharge) return;
            long total = checked(_retainedBytes + retainedLength - priorCharge);
            if (total > _maximumRetainedBytes) {
                throw PdfReadLimitException.Create(
                    PdfReadLimitKind.DecodedStreamBytes,
                    _maximumRetainedBytes,
                    total);
            }
            _charges[key] = retainedLength;
            _retainedBytes = total;
        }
    }
}

internal enum PdfIccProfileCacheRepresentation {
    ParsedProfile,
    DecodedBytes
}

internal static class PdfIccProfileCache {
    private static readonly ConditionalWeakTable<PdfStream, CacheSlot> Entries = new ConditionalWeakTable<PdfStream, CacheSlot>();

    internal static bool TryRead(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        out OfficeIccColorProfile? profile) =>
        TryRead(stream, objects, maxDecodedBytes, retentionBudget: null, out profile);

    internal static bool TryRead(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        PdfIccProfileRetentionBudget? retentionBudget,
        out OfficeIccColorProfile? profile) {
        ProfileCacheEntry entry = GetProfileEntry(stream, objects, maxDecodedBytes);
        if (entry.Profile != null) retentionBudget?.Charge(
            stream,
            PdfIccProfileCacheRepresentation.ParsedProfile,
            entry.RetainedLength);
        profile = entry.Profile;
        return profile != null;
    }

    internal static bool TryReadBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        out byte[] bytes) =>
        TryReadBytes(stream, objects, maxDecodedBytes, retentionBudget: null, out bytes);

    internal static bool TryReadBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes,
        PdfIccProfileRetentionBudget? retentionBudget,
        out byte[] bytes) {
        BytesCacheEntry entry = GetBytesEntry(stream, objects, maxDecodedBytes);
        if (entry.Decoded) retentionBudget?.Charge(
            stream,
            PdfIccProfileCacheRepresentation.DecodedBytes,
            entry.Bytes.LongLength);
        bytes = entry.Bytes;
        return entry.Decoded;
    }

    private static ProfileCacheEntry GetProfileEntry(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        CacheSlot slot = Entries.GetValue(stream, _ => new CacheSlot());
        lock (slot.Sync) {
            ProfileCacheEntry? entry = slot.ProfileEntry;
            if (entry == null) {
                entry = DecodeProfile(stream, objects, maxDecodedBytes);
                slot.ProfileEntry = entry;
            }
            EnsureLimit(entry.DecodedLength, maxDecodedBytes);
            return entry;
        }
    }

    private static BytesCacheEntry GetBytesEntry(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        CacheSlot slot = Entries.GetValue(stream, _ => new CacheSlot());
        lock (slot.Sync) {
            BytesCacheEntry? entry = slot.BytesEntry;
            if (entry == null) {
                entry = DecodeBytes(stream, objects, maxDecodedBytes);
                slot.BytesEntry = entry;
            }
            EnsureLimit(entry.Bytes.LongLength, maxDecodedBytes);
            return entry;
        }
    }

    private static ProfileCacheEntry DecodeProfile(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) {
        if (!PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] bytes, maxDecodedBytes)) {
            return new ProfileCacheEntry(decodedLength: 0L, retainedLength: 0L, null);
        }
        OfficeIccColorProfile.TryCreate(bytes, out OfficeIccColorProfile? profile);
        long retainedLength = profile == null
            ? bytes.LongLength
            : Math.Max(bytes.LongLength, profile.RetainedByteCount);
        return new ProfileCacheEntry(bytes.LongLength, retainedLength, profile);
    }

    private static BytesCacheEntry DecodeBytes(
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        int maxDecodedBytes) =>
        PdfImageStreamDecoder.TryDecode(stream, objects, out byte[] bytes, maxDecodedBytes)
            ? new BytesCacheEntry(bytes, decoded: true)
            : new BytesCacheEntry(Array.Empty<byte>(), decoded: false);

    private static void EnsureLimit(long decodedLength, int maxDecodedBytes) {
        if (decodedLength > maxDecodedBytes) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.DecodedStreamBytes,
                maxDecodedBytes,
                decodedLength);
        }
    }

    private sealed class ProfileCacheEntry {
        internal ProfileCacheEntry(long decodedLength, long retainedLength, OfficeIccColorProfile? profile) {
            DecodedLength = decodedLength;
            RetainedLength = retainedLength;
            Profile = profile;
        }
        internal long DecodedLength { get; }
        internal long RetainedLength { get; }
        internal OfficeIccColorProfile? Profile { get; }
    }

    private sealed class BytesCacheEntry {
        internal BytesCacheEntry(byte[] bytes, bool decoded) {
            Bytes = bytes;
            Decoded = decoded;
        }
        internal byte[] Bytes { get; }
        internal bool Decoded { get; }
    }

    private sealed class CacheSlot {
        internal object Sync { get; } = new object();
        internal ProfileCacheEntry? ProfileEntry { get; set; }
        internal BytesCacheEntry? BytesEntry { get; set; }
    }
}
