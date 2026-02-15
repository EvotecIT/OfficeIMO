using System.IO.Compression;

namespace OfficeIMO.Pdf.Filters;

internal static class FlateDecoder {
    public static byte[] Decode(byte[] data) {
        // Try zlib (RFC1950) first when available in this target
#if NET6_0_OR_GREATER
        if (TryZlib(data, out var result)) return result!;
#endif
        // Try raw Deflate
        if (TryInflate(data, out var result2)) return result2!;
        // Try skip zlib header (2 bytes) with raw Deflate
        if (data.Length > 2 && IsLikelyZlib(data)) {
            var sliced = new byte[data.Length - 2];
            Buffer.BlockCopy(data, 2, sliced, 0, sliced.Length);
            if (TryInflate(sliced, out var result3)) return result3!;
        }
        // Fallback to original
        return data;
    }

    private static bool TryInflate(byte[] input, out byte[]? output) {
        try {
            using var msIn = new MemoryStream(input);
            using var ds = new DeflateStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            using var msOut = new MemoryStream();
            ds.CopyTo(msOut);
            output = msOut.ToArray();
            return true;
        } catch { output = null; return false; }
    }

#if NET6_0_OR_GREATER
    private static bool TryZlib(byte[] input, out byte[]? output) {
        try {
            using var msIn = new MemoryStream(input);
            using var zs = new ZLibStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            using var msOut = new MemoryStream();
            zs.CopyTo(msOut);
            output = msOut.ToArray();
            return true;
        } catch { output = null; return false; }
    }
#endif

    private static bool IsLikelyZlib(byte[] d) {
        // RFC1950: first byte CMF low 4 bits = 8 for deflate; checksum of first two bytes mod 31 == 0
        if (d.Length < 2) return false;
        bool deflate = (d[0] & 0x0F) == 8;
        int cmfcm = (d[0] << 8) + d[1];
        bool mod = (cmfcm % 31) == 0;
        return deflate && mod;
    }
}

