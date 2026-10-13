using System;
using System.IO;

namespace OfficeIMO.Core.Internal;

internal static partial class OfficeArchiveSafety {
    // Entry extra fields select classic or ZIP64 sizes; the signature remains optional.
    // Accept only a bounded form that agrees with the authoritative central values.
    private static void ValidateZipDataDescriptor(Stream source, long offset, long bound, uint checksum, long compressed, long expanded, bool zip64) {
        if (offset > bound - 12) throw new InvalidDataException("ZIP data descriptor is missing or truncated.");
        var bytes = new byte[(int)Math.Min(24, bound - offset)];
        ReadExactlyAt(source, offset, bytes, bytes.Length);
        bool Matches(int start) {
            if (bytes.Length - start < (zip64 ? 20 : 12) || ReadZipUInt32(bytes, start) != checksum) return false;
            if (zip64) return ReadZipUInt64(bytes, start + 4) == (ulong)compressed && ReadZipUInt64(bytes, start + 12) == (ulong)expanded;
            return (ulong)ReadZipUInt32(bytes, start + 4) == (ulong)compressed && (ulong)ReadZipUInt32(bytes, start + 8) == (ulong)expanded;
        }
        if (Matches(0) || ReadZipUInt32(bytes, 0) == 0x08074b50U && Matches(4)) return;
        throw new InvalidDataException("ZIP data descriptor disagrees with its central record.");
    }
}
