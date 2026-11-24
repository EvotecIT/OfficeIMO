namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    // Entry extra fields select classic or ZIP64 sizes; the signature remains optional. Some writers use the
    // ZIP64 layout for members that need no ZIP64 sizes; that form is accepted only when both sizes agree.
    // Accept only a bounded form that agrees with the authoritative central values.
    private static void ValidateZipDataDescriptor(Stream source, long offset, long bound, uint checksum, long compressed, long expanded, bool zip64) {
        if (offset > bound - 12) throw new InvalidDataException("ZIP data descriptor is missing or truncated.");
        var bytes = new byte[(int)Math.Min(24, bound - offset)];
        ReadExactlyAt(source, offset, bytes, bytes.Length);
        bool Matches(int start, bool wide) {
            if (bytes.Length - start < (wide ? 20 : 12) || ReadZipUInt32(bytes, start) != checksum) return false;
            if (wide) return ReadZipUInt64(bytes, start + 4) == (ulong)compressed && ReadZipUInt64(bytes, start + 12) == (ulong)expanded;
            return (ulong)ReadZipUInt32(bytes, start + 4) == (ulong)compressed && (ulong)ReadZipUInt32(bytes, start + 8) == (ulong)expanded;
        }
        bool MatchesAt(int start) => Matches(start, zip64) || !zip64 && Matches(start, wide: true);
        if (MatchesAt(0) || ReadZipUInt32(bytes, 0) == 0x08074b50U && MatchesAt(4)) return;
        throw new InvalidDataException("ZIP data descriptor disagrees with its central record.");
    }
}
