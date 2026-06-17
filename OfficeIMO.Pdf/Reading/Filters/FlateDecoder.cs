using System;
using System.IO;
using System.IO.Compression;

namespace OfficeIMO.Pdf.Filters;

internal static class FlateDecoder {
    public static byte[] Decode(byte[] data) {
        // Try zlib (RFC1950) first when available in this target
#if NET6_0_OR_GREATER
        if (TryZlib(data, maxOutputBytes: null, out var result)) return result!;
#endif
        // Try raw Deflate
        if (TryInflate(data, maxOutputBytes: null, out var result2)) return result2!;
        // Try skip zlib header (2 bytes) with raw Deflate
        if (data.Length > 2 && IsLikelyZlib(data)) {
            var sliced = new byte[data.Length - 2];
            Buffer.BlockCopy(data, 2, sliced, 0, sliced.Length);
            if (TryInflate(sliced, maxOutputBytes: null, out var result3)) return result3!;
        }
        // Fallback to original
        return data;
    }

    public static bool TryDecode(byte[] data, int maxOutputBytes, out byte[] output) {
        if (maxOutputBytes < 0) {
            output = Array.Empty<byte>();
            return false;
        }

#if NET6_0_OR_GREATER
        if (TryZlib(data, maxOutputBytes, out var result)) {
            output = result!;
            return true;
        }
#endif

        if (TryInflate(data, maxOutputBytes, out var result2)) {
            output = result2!;
            return true;
        }

        if (data.Length > 2 && IsLikelyZlib(data)) {
            var sliced = new byte[data.Length - 2];
            Buffer.BlockCopy(data, 2, sliced, 0, sliced.Length);
            if (TryInflate(sliced, maxOutputBytes, out var result3)) {
                output = result3!;
                return true;
            }
        }

        if (data.Length <= maxOutputBytes) {
            output = data;
            return true;
        }

        output = Array.Empty<byte>();
        return false;
    }

    private static bool TryInflate(byte[] input, int? maxOutputBytes, out byte[]? output) {
        try {
            using var msIn = new MemoryStream(input);
            using var ds = new DeflateStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            return TryCopyToByteArray(ds, maxOutputBytes, out output);
        } catch { output = null; return false; }
    }

#if NET6_0_OR_GREATER
    private static bool TryZlib(byte[] input, int? maxOutputBytes, out byte[]? output) {
        try {
            using var msIn = new MemoryStream(input);
            using var zs = new ZLibStream(msIn, CompressionMode.Decompress, leaveOpen: true);
            return TryCopyToByteArray(zs, maxOutputBytes, out output);
        } catch { output = null; return false; }
    }
#endif

    private static bool TryCopyToByteArray(Stream source, int? maxOutputBytes, out byte[]? output) {
        using var msOut = new MemoryStream();
        var buffer = new byte[81920];
        int read;
        while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
            if (maxOutputBytes.HasValue && msOut.Length + read > maxOutputBytes.Value) {
                output = null;
                return false;
            }

            msOut.Write(buffer, 0, read);
        }

        output = msOut.ToArray();
        return true;
    }


    private static bool IsLikelyZlib(byte[] d) {
        // RFC1950: first byte CMF low 4 bits = 8 for deflate; checksum of first two bytes mod 31 == 0
        if (d.Length < 2) return false;
        bool deflate = (d[0] & 0x0F) == 8;
        int cmfcm = (d[0] << 8) + d[1];
        bool mod = (cmfcm % 31) == 0;
        return deflate && mod;
    }
}

